from ..config import LINE_LENGTH
from ..utils import format_date


def generate_spool_line(row_data, header_values, invoice_data, is_spare):
    line = [' '] * LINE_LENGTH

    row_data = {k: str(v) if v is not None else '' for k, v in row_data.items()}

    def set_field(start, end, value, align='left'):
        width = end - start
        value_str = str(value) if value else ''
        if align == 'left':
            formatted = value_str[:width].ljust(width)
        else:
            formatted = value_str[:width].rjust(width)
        for i, char in enumerate(formatted):
            if start + i < LINE_LENGTH:
                line[start + i] = char

    invoice_date_raw = header_values.get('invoice_date')
    challan_date_raw = header_values.get('challan_date') or invoice_date_raw

    challan_date = format_date(challan_date_raw, '%d-%b-%Y')
    invoice_date = format_date(invoice_date_raw, '%d-%b-%Y', upper=True)

    set_field(0, 4, header_values.get('vendor_code') or 'X539')
    set_field(4, 20, header_values.get('challan_no') or header_values.get('invoice_no'))
    set_field(20, 31, challan_date)
    set_field(31, 47, header_values.get('invoice_no'))
    set_field(47, 71, invoice_date)

    oe_prefix = '1' if not is_spare else 'S'
    set_field(82, 83, oe_prefix)
    set_field(83, 98, row_data.get('schedule_no', ''))
    set_field(98, 113, row_data.get('item_code', ''))
    set_field(113, 125, row_data.get('qty', ''))
    set_field(125, 138, header_values.get('po_number'))

    bin_qty_val = str(row_data.get('bin_qty', '')).strip()
    bin_qty_field = f"    {bin_qty_val}" if bin_qty_val else ''
    set_field(138, 150, bin_qty_field)

    if is_spare and row_data.get('batch_no'):
        set_field(194, 204, str(row_data['batch_no']).strip())

    set_field(204, 219, row_data.get('gst_no', ''))
    set_field(219, 227, row_data.get('hsn_code', ''))

    cgst = str(row_data.get('cgst_amt')) if row_data.get('cgst_amt') else '0'
    sgst = str(row_data.get('sgst_amt')) if row_data.get('sgst_amt') else '0'
    cgst_portion = cgst.ljust(16) + sgst[:2]
    set_field(227, 245, cgst_portion)
    sgst_portion = sgst[2:] if len(sgst) > 2 else ''
    set_field(245, 265, sgst_portion)

    set_field(275, 276, str(row_data.get('eway_bill')) if row_data.get('eway_bill') else '0')
    set_field(276, 290, row_data.get('igst_amt', ''))
    set_field(290, 354, invoice_data.get('irn_number', ''))
    set_field(354, 366, row_data.get('basic_price', ''))
    set_field(366, 390, row_data.get('total_value', ''))

    return ''.join(line)
