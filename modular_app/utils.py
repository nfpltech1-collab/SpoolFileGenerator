import os
from datetime import datetime

from .config import DATE_INPUT_FORMATS


def normalize_item_code(code):
    if not code:
        return ''
    return str(code).replace('-', '').replace(' ', '').upper()


def format_date(value, output_format, input_formats=None, upper=False):
    if not value:
        return value
    formats = input_formats or DATE_INPUT_FORMATS
    for fmt in formats:
        try:
            dt = datetime.strptime(value, fmt)
            formatted = dt.strftime(output_format)
            return formatted.upper() if upper else formatted
        except Exception:
            continue
    return value


def get_initial_dir(last_path):
    if last_path and os.path.exists(os.path.dirname(last_path)):
        return os.path.dirname(last_path)
    return os.path.expanduser("~")
