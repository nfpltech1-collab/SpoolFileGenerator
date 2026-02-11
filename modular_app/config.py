COLUMNS = [
    ("unload_no", "Unload No", 100),
    ("schedule_no", "Schedule No", 140),
    ("item_code", "Item Code", 100),
    ("qty", "Qty", 50),
    ("po_number", "PO Number", 80),
    ("f57_2no", "57F2 No", 60),
    ("bin_qty", "Bin Qty", 60),
    ("remarks", "Remarks", 60),
    ("batch_no", "Batch No", 60),
    ("location", "Location", 60),
    ("gst_no", "GST No", 130),
    ("hsn_code", "HSN Code", 80),
    ("cgst_amt", "CGST Amt", 90),
    ("sgst_amt", "SGST Amt", 90),
    ("igst_amt", "IGST Amt", 70),
    ("eway_bill", "E-Way Bill", 70),
    ("basic_price", "Basic Price", 90),
    ("total_value", "Total Invoice", 100),
    ("tool_amort", "Tool/Amort", 70),
]

LINE_LENGTH = 390

DATE_INPUT_FORMATS = [
    '%d-%b-%y',
    '%d-%b-%Y',
    '%d/%m/%Y',
    '%d-%m-%Y',
]

GSTIN_PATTERN = r'^[0-3][0-9][A-Z]{5}[0-9]{4}[A-Z][1-9A-Z][Z][0-9A-Z]$'
