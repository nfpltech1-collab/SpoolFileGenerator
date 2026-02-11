import re

from ..config import GSTIN_PATTERN


def validate_required_fields(header_values, invoice_data=None):
    required_fields = {
        'vendor_code': 'Vendor Code',
        'challan_no': 'Challan No',
        'challan_date': 'Challan Date',
        'invoice_no': 'Invoice No',
        'invoice_date': 'Invoice Date',
        'po_number': 'PO Number',
    }

    missing = []
    invoice_data = invoice_data or {}

    for field_id, field_name in required_fields.items():
        value = header_values.get(field_id, '') or invoice_data.get(field_id, '')
        if not value or not str(value).strip():
            missing.append(field_name)

    if missing:
        return False, f"Missing required fields:\n• " + "\n• ".join(missing)
    return True, ""


def validate_preview_rows(preview_data, is_spare):
    common_required = [
        'unload_no',
        'schedule_no',
        'item_code',
        'qty',
        'po_number',
        'gst_no',
        'hsn_code',
        'cgst_amt',
        'sgst_amt',
        'basic_price',
        'total_value',
    ]

    oe_required = common_required + ['bin_qty']
    spare_required = common_required + ['batch_no', 'bin_qty']
    required_fields = spare_required if is_spare else oe_required

    field_names = {
        'unload_no': 'Unload No',
        'schedule_no': 'Schedule No (KANBAN)',
        'item_code': 'Item Code (Part No)',
        'qty': 'Qty',
        'po_number': 'PO Number',
        'bin_qty': 'Bin Qty',
        'batch_no': 'Batch No',
        'gst_no': 'GST No',
        'hsn_code': 'HSN Code',
        'cgst_amt': 'CGST Amount',
        'sgst_amt': 'SGST Amount',
        'basic_price': 'Basic Price',
        'total_value': 'Total Invoice Value',
    }

    errors = []

    for row_idx, row in enumerate(preview_data):
        row_errors = []
        for field_id in required_fields:
            value = row.get(field_id, '')
            if not value or not str(value).strip():
                row_errors.append(field_names.get(field_id, field_id))

        gst_no = row.get('gst_no', '')
        if gst_no and str(gst_no).strip():
            gst_no = str(gst_no).strip()
            if not re.match(GSTIN_PATTERN, gst_no, re.IGNORECASE):
                row_errors.append(f"GST No format invalid ({gst_no})")
            else:
                state_code = int(gst_no[:2])
                if state_code < 1 or state_code > 37:
                    row_errors.append(f"GST No state code invalid ({state_code})")

        if row_errors:
            item_code = row.get('item_code', f'Row {row_idx + 1}')
            errors.append(f"Row '{item_code}': {', '.join(row_errors)}")

    return len(errors) == 0, errors
