import tkinter as tk
from datetime import datetime
import os
import sys

if __package__ is None or __package__ == "":
    parent_dir = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, os.path.dirname(parent_dir))
    from modular_app.ui.view import SpoolAppView
    from modular_app.controller import SpoolAppController
else:
    from .ui.view import SpoolAppView
    from .controller import SpoolAppController


def main():
    root = tk.Tk()
    root.title("Spool File Generator (GST) v2")
    root.geometry("1400x800")
    root.minsize(1200, 700)

    callbacks = {
        'browse_invoices': lambda: controller.browse_invoices(),
        'browse_excel': lambda: controller.browse_excel(),
        'find_workbook': lambda: controller.find_workbook_by_date(),
        'load_preview': lambda: controller.load_preview(),
        'generate_all': lambda: controller.generate_all_spool(),
        'save_changes': lambda: controller.save_changes(),
        'clear_all': lambda: controller.clear_all(),
        'view_output': lambda: controller.view_output(),
        'prev_preview': lambda: controller.prev_preview(),
        'next_preview': lambda: controller.next_preview(),
        'close': lambda: controller.close(),
        'get_today': lambda: datetime.now().strftime('%d-%m-%Y'),
    }

    view = SpoolAppView(root, callbacks)
    controller = SpoolAppController(root, view)

    root.mainloop()


if __name__ == "__main__":
    main()
