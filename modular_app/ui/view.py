import tkinter as tk
from tkinter import ttk
import os
from PIL import Image, ImageTk  # Added for Nagarkot GUI standards

from ..config import COLUMNS


class SpoolAppView:
    def __init__(self, root, callbacks):
        self.root = root
        self.callbacks = callbacks

        self._setup_styles()
        self._create_widgets()

    def _setup_styles(self):
        """Configure ttk styles for Nagarkot Corporate Identity."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Corporate Colors
        NAGARKOT_BLUE = '#0056b3'
        WHITE = '#ffffff'
        LIGHT_GRAY = '#f8f9fa'
        TEXT_DARK = '#212529'
        GRAY_BORDER = '#ced4da'
        SUCCESS_GREEN = '#28a745'
        ERROR_RED = '#dc3545'
        
        # General App Style
        self.root.configure(bg=WHITE)
        style.configure('.', background=WHITE, foreground=TEXT_DARK, font=('Segoe UI', 10))
        style.configure('TFrame', background=WHITE)
        style.configure('TLabelframe', background=WHITE, bordercolor=GRAY_BORDER)
        style.configure('TLabelframe.Label', background=WHITE, foreground=NAGARKOT_BLUE, font=('Segoe UI', 10, 'bold'))
        
        # Header Styles
        style.configure("Header.TLabel", font=('Segoe UI', 11, 'bold'), foreground=NAGARKOT_BLUE, background=WHITE)
        style.configure("Title.TLabel", font=('Helvetica', 18, 'bold'), foreground=NAGARKOT_BLUE, background=WHITE)
        style.configure("Subtitle.TLabel", font=('Segoe UI', 10), foreground='#6c757d', background=WHITE)
        
        # Status Bar Styles
        style.configure("Status.TFrame", background=LIGHT_GRAY, relief=tk.GROOVE)
        style.configure("Copyright.TLabel", font=('Segoe UI', 8), foreground='#6c757d', background=LIGHT_GRAY)
        style.configure("Success.TLabel", font=('Segoe UI', 9, 'bold'), foreground=SUCCESS_GREEN, background=LIGHT_GRAY)
        style.configure("Error.TLabel", font=('Segoe UI', 9, 'bold'), foreground=ERROR_RED, background=LIGHT_GRAY)
        style.configure("Info.TLabel", font=('Segoe UI', 9), foreground=NAGARKOT_BLUE, background=LIGHT_GRAY)
        
        # Button Styles
        style.configure("Primary.TButton", 
                        font=('Segoe UI', 9, 'bold'), 
                        background=NAGARKOT_BLUE, 
                        foreground='white',
                        borderwidth=0,
                        focuscolor=NAGARKOT_BLUE,
                        padding=(15, 8))
        style.map("Primary.TButton", 
                  background=[('active', '#004494'), ('pressed', '#003366')])
        
        style.configure("Action.TButton", 
                        font=('Segoe UI', 9), 
                        background=WHITE, 
                        foreground=TEXT_DARK,
                        borderwidth=1,
                        bordercolor=GRAY_BORDER,
                        focuscolor=LIGHT_GRAY,
                        padding=(10, 5))
        style.map("Action.TButton", 
                  background=[('active', LIGHT_GRAY), ('pressed', '#e2e6ea')])

        # Treeview (Table) Style
        style.configure("Treeview", 
                        font=('Consolas', 9), 
                        rowheight=25,
                        background=WHITE, 
                        fieldbackground=WHITE,
                        foreground=TEXT_DARK)
        style.configure("Treeview.Heading", 
                        font=('Segoe UI', 9, 'bold'), 
                        background=NAGARKOT_BLUE, 
                        foreground='white',
                        relief='flat')
        style.map("Treeview.Heading", 
                  background=[('active', '#004494')])

    def _create_widgets(self):
        # === HEADER SECTION (Logo & Title) ===
        header_frame = ttk.Frame(self.root, padding="10 15 10 15")
        header_frame.pack(fill=tk.X)
        
        # Set minimum height for header frame
        header_frame.configure(height=60)
        
        # 1. Logo (packed to the left)
        try:
            # Locate logo in project root (3 levels up from this file)
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            logo_path = os.path.join(base_dir, "logo.png")
            
            if os.path.exists(logo_path):
                pil_image = Image.open(logo_path)
                # Resize to Height = 20px
                h_size = 20
                aspect = pil_image.width / pil_image.height
                w_size = int(h_size * aspect)
                pil_image = pil_image.resize((w_size, h_size), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(pil_image)
                
                logo_label = ttk.Label(header_frame, image=self.logo_img, background='#ffffff')
                logo_label.pack(side=tk.LEFT, padx=10, pady=5)
            else:
                ttk.Label(header_frame, text="[LOGO]", style="Header.TLabel").pack(side=tk.LEFT, padx=10, pady=5)
        except Exception as e:
            print(f"Logo load error: {e}")
            
        # 2. Title & Subtitle (centered across full window width)
        title_container = ttk.Frame(header_frame)
        title_container.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        title_label = ttk.Label(title_container, text="SPOOL FILE GENERATOR", style="Title.TLabel")
        title_label.pack(pady=(0, 2))
        
        subtitle_label = ttk.Label(title_container, text="Nagarkot Forwarders Pvt Ltd", style="Subtitle.TLabel")
        subtitle_label.pack()

        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(file_frame, text="Invoice PDF(s):", style="Header.TLabel").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.invoice_path_var = tk.StringVar()
        self.invoice_entry = ttk.Entry(file_frame, textvariable=self.invoice_path_var, width=60, state='readonly')
        self.invoice_entry.grid(row=0, column=1, padx=5, sticky=tk.EW)
        ttk.Button(file_frame, text="Browse...", command=self.callbacks['browse_invoices']).grid(row=0, column=2, padx=5)

        ttk.Label(file_frame, text="Excel/CSV:", style="Header.TLabel").grid(row=1, column=0, sticky=tk.W, padx=5, pady=(10, 0))
        self.excel_path_var = tk.StringVar()
        self.excel_entry = ttk.Entry(file_frame, textvariable=self.excel_path_var, width=60, state='readonly')
        self.excel_entry.grid(row=1, column=1, padx=5, pady=(10, 0), sticky=tk.EW)
        ttk.Button(file_frame, text="Browse...", command=self.callbacks['browse_excel']).grid(row=1, column=2, padx=5, pady=(10, 0))

        ttk.Label(file_frame, text="Dispatch Date:", style="Header.TLabel").grid(row=2, column=0, sticky=tk.W, padx=5, pady=(10, 0))

        date_frame = ttk.Frame(file_frame)
        date_frame.grid(row=2, column=1, padx=5, pady=(10, 0), sticky=tk.W)

        self.date_var = tk.StringVar(value=self.callbacks['get_today']())
        self.date_entry = ttk.Entry(date_frame, textvariable=self.date_var, width=12)
        self.date_entry.pack(side=tk.LEFT, padx=2)

        ttk.Label(date_frame, text="(DD-MM-YYYY)").pack(side=tk.LEFT, padx=5)
        ttk.Button(date_frame, text="Find Workbook", command=self.callbacks['find_workbook']).pack(side=tk.LEFT, padx=10)

        file_frame.columnconfigure(1, weight=1)

        header_frame = ttk.LabelFrame(main_frame, text="Invoice Header", padding="10")
        header_frame.pack(fill=tk.X, pady=(0, 10))

        self.header_entries = {}

        ttk.Label(header_frame, text="Vendor Code:", style="Header.TLabel").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.header_entries['vendor_code'] = ttk.Entry(header_frame, width=15)
        self.header_entries['vendor_code'].grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Label(header_frame, text="Challan No:", style="Header.TLabel").grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        self.header_entries['challan_no'] = ttk.Entry(header_frame, width=20)
        self.header_entries['challan_no'].grid(row=0, column=3, sticky=tk.W, padx=5, pady=2)

        ttk.Label(header_frame, text="Challan Date:", style="Header.TLabel").grid(row=0, column=4, sticky=tk.W, padx=5, pady=2)
        self.header_entries['challan_date'] = ttk.Entry(header_frame, width=15)
        self.header_entries['challan_date'].grid(row=0, column=5, sticky=tk.W, padx=5, pady=2)

        ttk.Label(header_frame, text="Invoice No:", style="Header.TLabel").grid(row=0, column=6, sticky=tk.W, padx=5, pady=2)
        self.header_entries['invoice_no'] = ttk.Entry(header_frame, width=20)
        self.header_entries['invoice_no'].grid(row=0, column=7, sticky=tk.W, padx=5, pady=2)

        ttk.Label(header_frame, text="Invoice Date:", style="Header.TLabel").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.header_entries['invoice_date'] = ttk.Entry(header_frame, width=15)
        self.header_entries['invoice_date'].grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Label(header_frame, text="PO Number:", style="Header.TLabel").grid(row=1, column=2, sticky=tk.W, padx=5, pady=2)
        self.header_entries['po_number'] = ttk.Entry(header_frame, width=15)
        self.header_entries['po_number'].grid(row=1, column=3, sticky=tk.W, padx=5, pady=2)

        ttk.Label(header_frame, text="Excise Amount:", style="Header.TLabel").grid(row=1, column=4, sticky=tk.W, padx=5, pady=2)
        self.header_entries['excise_amt'] = ttk.Entry(header_frame, width=15)
        self.header_entries['excise_amt'].grid(row=1, column=5, sticky=tk.W, padx=5, pady=2)

        ttk.Label(header_frame, text="Sales Tax:", style="Header.TLabel").grid(row=1, column=6, sticky=tk.W, padx=5, pady=2)
        self.header_entries['sales_tax'] = ttk.Entry(header_frame, width=15)
        self.header_entries['sales_tax'].grid(row=1, column=7, sticky=tk.W, padx=5, pady=2)

        ttk.Label(header_frame, text="OE/Spares:", style="Header.TLabel").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)

        self.oe_spares_var = tk.StringVar(value="OE")
        radio_frame = ttk.Frame(header_frame)
        radio_frame.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Radiobutton(radio_frame, text="O/E", variable=self.oe_spares_var, value="OE").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(radio_frame, text="Spare", variable=self.oe_spares_var, value="Spare").pack(side=tk.LEFT, padx=5)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        buttons = [
            ("Load Preview", self.callbacks['load_preview'], "Primary.TButton"),
            ("Generate All", self.callbacks['generate_all'], "Primary.TButton"),
            ("Save", self.callbacks['save_changes'], "Action.TButton"),
            ("Clear", self.callbacks['clear_all'], "Action.TButton"),
            ("View Output", self.callbacks['view_output'], "Action.TButton"),
            ("Close", self.callbacks['close'], "Action.TButton"),
        ]

        for text, command, style in buttons:
            ttk.Button(button_frame, text=text, command=command, style=style).pack(side=tk.LEFT, padx=5)

        table_frame = ttk.LabelFrame(main_frame, text="Preview Table (Double-click to edit)", padding="5")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        nav_frame = ttk.Frame(table_frame)
        nav_frame.pack(fill=tk.X, pady=(0, 5))

        self.prev_btn = ttk.Button(nav_frame, text="◀ Previous", command=self.callbacks['prev_preview'], state='disabled')
        self.prev_btn.pack(side=tk.LEFT, padx=5)

        self.preview_label_var = tk.StringVar(value="No invoices loaded")
        ttk.Label(nav_frame, textvariable=self.preview_label_var, font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=20)

        self.next_btn = ttk.Button(nav_frame, text="Next ▶", command=self.callbacks['next_preview'], state='disabled')
        self.next_btn.pack(side=tk.LEFT, padx=5)

        tree_container = ttk.Frame(table_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        h_scroll = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        v_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        columns = [col[0] for col in COLUMNS]
        self.tree = ttk.Treeview(tree_container, columns=columns, show='headings',
                                  xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        self.tree.pack(fill=tk.BOTH, expand=True)

        h_scroll.config(command=self.tree.xview)
        v_scroll.config(command=self.tree.yview)

        for col_id, col_name, col_width in COLUMNS:
            self.tree.heading(col_id, text=col_name)
            self.tree.column(col_id, width=col_width, minwidth=50)

        self.tree.bind('<Double-1>', self._on_cell_double_click)

        # === FOOTER / STATUS BAR ===
        footer_frame = ttk.Frame(self.root, style="Status.TFrame", padding=(10, 5))
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        ttk.Label(footer_frame, text="© Nagarkot Forwarders Pvt Ltd", style="Copyright.TLabel").pack(side=tk.LEFT)
        
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(footer_frame, textvariable=self.status_var, style="Info.TLabel")
        self.status_bar.pack(side=tk.RIGHT, padx=10)

    def _on_cell_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)

        if not row:
            return

        col_idx = int(column[1:]) - 1
        col_id = COLUMNS[col_idx][0]

        bbox = self.tree.bbox(row, column)
        if not bbox:
            return

        item = self.tree.item(row)
        current_value = item['values'][col_idx]

        entry = ttk.Entry(self.tree, width=bbox[2])
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus()

        def save_edit(event=None):
            new_value = entry.get()
            values = list(item['values'])
            values[col_idx] = new_value
            self.tree.item(row, values=values)
            entry.destroy()

        def cancel_edit(event=None):
            entry.destroy()

        entry.bind('<Return>', save_edit)
        entry.bind('<Escape>', cancel_edit)
        entry.bind('<FocusOut>', save_edit)

    def set_status(self, message, style='normal'):
        self.status_var.set(message)
        if style == 'success':
            self.status_bar.configure(style='Success.TLabel')
        elif style == 'error':
            self.status_bar.configure(style='Error.TLabel')
        elif style == 'info':
            self.status_bar.configure(style='Info.TLabel')
        else:
            self.status_bar.configure(style='TLabel')
        self.root.update()

    def set_invoice_path_display(self, text):
        self.invoice_path_var.set(text)

    def set_excel_path_display(self, text):
        self.excel_path_var.set(text)

    def get_dispatch_date(self):
        return self.date_var.get().strip()

    def set_preview_label(self, text):
        self.preview_label_var.set(text)

    def set_nav_state(self, prev_enabled, next_enabled):
        self.prev_btn.config(state='normal' if prev_enabled else 'disabled')
        self.next_btn.config(state='normal' if next_enabled else 'disabled')

    def get_oe_spares(self):
        return self.oe_spares_var.get()

    def get_header_values(self):
        return {field: entry.get().strip() for field, entry in self.header_entries.items()}

    def set_header_values(self, data):
        for field, entry in self.header_entries.items():
            entry.delete(0, tk.END)
            if field in data:
                entry.insert(0, data[field])

    def clear_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def set_table_rows(self, rows):
        self.clear_table()
        for row_data in rows:
            values = [row_data[col[0]] for col in COLUMNS]
            self.tree.insert('', 'end', values=values)

    def get_table_rows(self):
        rows = []
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id)['values']
            row_data = {}
            for col_idx, (col_id, _, _) in enumerate(COLUMNS):
                row_data[col_id] = values[col_idx]
            rows.append(row_data)
        return rows
