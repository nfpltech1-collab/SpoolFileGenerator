import os
import re

import pdfplumber

try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

from ..config import GSTIN_PATTERN
from ..utils import normalize_item_code


def extract_invoice_data(invoice_path):
    invoice_data = {}
    invoice_line_items = {}
    validation_info = {}

    with pdfplumber.open(invoice_path) as pdf:
        full_text = ""
        has_images = False

        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"

            if page.images:
                has_images = True

        validation_info['full_text'] = full_text
        validation_info['has_images'] = has_images
        validation_info['pdf_path'] = invoice_path

        inv_match = re.search(r'Invoice\s+Number\s*:\s*([A-Z0-9/\-]+)', full_text, re.IGNORECASE)
        if inv_match:
            invoice_data['invoice_no'] = inv_match.group(1).strip()

        date_match = re.search(r'Invoice\s+Date\s*:\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})', full_text, re.IGNORECASE)
        if date_match:
            invoice_data['invoice_date'] = date_match.group(1).strip()

        po_match = re.search(r'Cust\s+PO\s+No\.?\s*:\s*(\d+)', full_text, re.IGNORECASE)
        if po_match:
            invoice_data['po_number'] = po_match.group(1).strip()

        ref_match = re.search(r'Reference\s+No\.?\s*:\s*([A-Z]\d{3,})', full_text, re.IGNORECASE)
        if ref_match:
            invoice_data['vendor_code'] = ref_match.group(1).strip()

        gst_match = re.search(r'GSTIN\s+Number\s*:\s*(\d{2}[A-Z]{5}\d{4}[A-Z]\d[A-Z\d]{2})', full_text, re.IGNORECASE)
        if gst_match:
            invoice_data['gst_no'] = gst_match.group(1).strip()

        irn_match = re.search(r'IRN\s*(?:NO)?[:\s]*([a-f0-9]{64})', full_text, re.IGNORECASE)
        if irn_match:
            invoice_data['irn_number'] = irn_match.group(1).strip()

        total_match = re.search(r'Invoice\s+Amount\s*\(INR\)\s*([\d,]+\.?\d*)', full_text, re.IGNORECASE)
        if total_match:
            invoice_data['total_value'] = total_match.group(1).replace(',', '').strip()

        tax_totals = re.search(
            r'^([\d,]{6,}\.[0-9]{2})\s+([\d,]+\.[0-9]{2})\s+([\d,]+\.[0-9]{2})\s+([\d,]+\.[0-9]{2})\s+[\d,]+\.[0-9]{2}\s*$',
            full_text, re.MULTILINE
        )
        if tax_totals:
            invoice_data['cgst_amt'] = tax_totals.group(2).replace(',', '')
            invoice_data['sgst_amt'] = tax_totals.group(3).replace(',', '')
            igst_val = tax_totals.group(4).replace(',', '')
            invoice_data['igst_amt'] = '' if float(igst_val) == 0 else igst_val

        line_pattern = re.compile(
            r'(\d)\s+(\d{6}-\d{5})\s+(\d{8})\s+(\d+)\.00\s+Nos\s+([\d,]+\.\d{3})',
            re.IGNORECASE
        )
        item_code_pattern = re.compile(
            r'\(\s*([0-9A-Z]{6}-?[0-9A-Z]{5}(?:-\d+)?)\s*\)',
            re.IGNORECASE
        )

        line_matches = list(line_pattern.finditer(full_text))
        for i, match in enumerate(line_matches):
            sno = match.group(1)
            material_code = match.group(2)
            hsn_code = match.group(3)
            qty = match.group(4)
            rate = match.group(5).replace(',', '')

            search_start = match.end()
            search_end = line_matches[i + 1].start() if i + 1 < len(line_matches) else len(full_text)
            search_text = full_text[search_start:search_end]
            item_code_match = item_code_pattern.search(search_text)

            if item_code_match:
                item_code_raw = item_code_match.group(1)
                item_code_norm = normalize_item_code(item_code_raw)
                invoice_line_items[item_code_norm] = {
                    'sno': sno,
                    'material_code': material_code,
                    'item_code': item_code_raw,
                    'hsn_code': hsn_code,
                    'qty': qty,
                    'rate': rate,
                }

        invoice_data['eway_bill'] = '0'

    return invoice_data, invoice_line_items, validation_info


def get_invoice_item(part_number, invoice_line_items):
    if not part_number:
        return None

    part_norm = normalize_item_code(part_number)

    if part_norm in invoice_line_items:
        return invoice_line_items[part_norm]

    for key, value in invoice_line_items.items():
        if key in part_norm or part_norm in key:
            return value

    return None


def validate_invoice_integrity(invoice_data, validation_info):
    errors = []
    full_text = validation_info.get('full_text', '')
    pdf_path = validation_info.get('pdf_path', '')

    irn = invoice_data.get('irn_number', '')
    if not irn:
        errors.append("IRN Number not found in invoice")
    elif len(irn) != 64:
        errors.append(f"IRN Number invalid length: {len(irn)} chars (expected 64)")
    elif not re.match(r'^[a-f0-9]{64}$', irn, re.IGNORECASE):
        errors.append("IRN Number contains invalid characters (must be alphanumeric hex)")

    original_patterns = [
        r'Original\s+for\s*\n?\s*Recipient',
        r'Original\s+for\s+Recipient',
        r'Tax\s+Invoice\s+Original',
    ]
    is_original = any(re.search(pattern, full_text, re.IGNORECASE) for pattern in original_patterns)
    if not is_original:
        errors.append("Invoice is not 'Original for Recipient' copy")

    has_digital_signature = _check_digital_signature(pdf_path, full_text)
    if not has_digital_signature:
        errors.append("Digital Signature not found (must have 'Digitally signed by...' with signer name)")

    gst_no = invoice_data.get('gst_no', '')
    if not gst_no:
        errors.append("GSTIN Number not found in invoice")
    else:
        if not re.match(GSTIN_PATTERN, gst_no, re.IGNORECASE):
            errors.append(f"GSTIN Number format invalid: {gst_no}")
        else:
            state_code = int(gst_no[:2])
            if state_code < 1 or state_code > 37:
                errors.append(f"GSTIN state code invalid: {state_code} (must be 01-37)")

    return len(errors) == 0, errors


def _check_digital_signature(pdf_path, pdfplumber_text):
    digital_sign_patterns = [
        r'Digitally\s+signed\s+by\s+[A-Z\s]+',
        r'Digitally\s+signed\s+by.*Date:\d{4}\.\d{2}\.\d{2}',
        r'Digital\s+Signature.*Date:',
    ]
    for pattern in digital_sign_patterns:
        if re.search(pattern, pdfplumber_text, re.IGNORECASE):
            return True

    if HAS_PYMUPDF and pdf_path and os.path.exists(pdf_path):
        try:
            doc = fitz.open(pdf_path)
            for page in doc:
                widgets = list(page.widgets())
                for widget in widgets:
                    if widget.field_type_string == 'Signature':
                        doc.close()
                        return True

                text_dict = page.get_text("dict")
                all_text = []
                for block in text_dict.get("blocks", []):
                    if "lines" in block:
                        for line in block["lines"]:
                            for span in line["spans"]:
                                all_text.append(span.get("text", ""))

                page_text = " ".join(all_text)
                if any(re.search(pattern, page_text, re.IGNORECASE) for pattern in digital_sign_patterns):
                    doc.close()
                    return True
            doc.close()
        except Exception:
            pass

    if pdf_path and os.path.exists(pdf_path):
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    annots = page.annots
                    if annots:
                        for annot in annots:
                            annot_data = annot.get('data', {})
                            if annot_data.get('FT') == '/Sig':
                                sig_value = annot_data.get('V', {})
                                if isinstance(sig_value, dict):
                                    signer_name = sig_value.get('Name', b'')
                                    if isinstance(signer_name, bytes):
                                        signer_name = signer_name.decode('utf-8', errors='ignore')
                                    if signer_name:
                                        return True
        except Exception:
            pass

    return False
