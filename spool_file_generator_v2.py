"""
Spool File Generator (GST) - Version 2
Redesigned GUI with editable preview table and file browse options.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk  # Added for Nagarkot GUI standards
import pandas as pd
import pdfplumber
import re
import os
import sys
from datetime import datetime, timedelta

# Try to import PyMuPDF for better digital signature detection
try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False


class SpoolFileGeneratorV2:
    """Main application class for generating spool files with modern GUI."""
    
    # Table column definitions (19 columns)
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
    
    def __init__(self, root):
        self.root = root
        self.root.title("Spool File Generator (GST) v2")
        self.root.geometry("1400x800")
        self.root.minsize(1200, 700)
        
        # Configure style
        self.setup_styles()
        
        # Directories
        # Directories
        if getattr(sys, 'frozen', False):
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_dir = os.path.join(self.base_dir, "SpoolOutput")
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Data storage
        self.invoice_path = None
        self.invoice_paths = []  # For bulk processing - multiple invoices
        self.excel_path = None
        self.invoice_data = {}
        self.invoice_line_items = {}
        self.preview_data = []
        self.selected_date = None  # User-selected date for workbook search
        
        # Multi-invoice preview storage
        self.all_previews = []  # List of (invoice_data, header_data, preview_data) tuples
        self.current_preview_index = 0
        
        self.create_widgets()
    
    def setup_styles(self):
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
        # Primary: Blue bg, White text
        style.configure("Primary.TButton", 
                        font=('Segoe UI', 9, 'bold'), 
                        background=NAGARKOT_BLUE, 
                        foreground='white',
                        borderwidth=0,
                        focuscolor=NAGARKOT_BLUE,
                        padding=(15, 8))
        style.map("Primary.TButton", 
                  background=[('active', '#004494'), ('pressed', '#003366')])
        
        # Action/Secondary: White bg, Gray border
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
    
    def create_widgets(self):
        """Create all GUI widgets following Nagarkot standards."""
        # === HEADER SECTION (Logo & Title) ===
        header_frame = ttk.Frame(self.root, padding="10 15 10 15")
        header_frame.pack(fill=tk.X)
        
        # Set minimum height for header frame
        header_frame.configure(height=60)
        
        # 1. Logo (packed to the left)
        try:
            # Look for logo in current directory
            logo_path = os.path.join(self.base_dir, "logo.png")
            if os.path.exists(logo_path):
                pil_image = Image.open(logo_path)
                # Resize to Height = 20px (maintain aspect ratio)
                h_size = 20
                aspect = pil_image.width / pil_image.height
                w_size = int(h_size * aspect)
                pil_image = pil_image.resize((w_size, h_size), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(pil_image)
                
                logo_label = ttk.Label(header_frame, image=self.logo_img, background='#ffffff')
                logo_label.pack(side=tk.LEFT, padx=10, pady=5)
            else:
                # Fallback text if logo missing
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
        
        # === FILE SELECTION SECTION ===
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Invoice PDF row
        ttk.Label(file_frame, text="Invoice PDF(s):", style="Header.TLabel").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.invoice_path_var = tk.StringVar()
        self.invoice_entry = ttk.Entry(file_frame, textvariable=self.invoice_path_var, width=60, state='readonly')
        self.invoice_entry.grid(row=0, column=1, padx=5, sticky=tk.EW)
        ttk.Button(file_frame, text="Browse...", command=self.browse_invoices).grid(row=0, column=2, padx=5)
        
        # Excel/CSV row
        ttk.Label(file_frame, text="Excel/CSV:", style="Header.TLabel").grid(row=1, column=0, sticky=tk.W, padx=5, pady=(10, 0))
        self.excel_path_var = tk.StringVar()
        self.excel_entry = ttk.Entry(file_frame, textvariable=self.excel_path_var, width=60, state='readonly')
        self.excel_entry.grid(row=1, column=1, padx=5, pady=(10, 0), sticky=tk.EW)
        ttk.Button(file_frame, text="Browse...", command=self.browse_excel).grid(row=1, column=2, padx=5, pady=(10, 0))
        
        # Date selection row (below Excel/CSV)
        ttk.Label(file_frame, text="Dispatch Date:", style="Header.TLabel").grid(row=2, column=0, sticky=tk.W, padx=5, pady=(10, 0))
        
        date_frame = ttk.Frame(file_frame)
        date_frame.grid(row=2, column=1, padx=5, pady=(10, 0), sticky=tk.W)
        
        # Date entry field (DD-MM-YYYY format)
        self.date_var = tk.StringVar(value=datetime.now().strftime('%d-%m-%Y'))
        self.date_entry = ttk.Entry(date_frame, textvariable=self.date_var, width=12)
        self.date_entry.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(date_frame, text="(DD-MM-YYYY)").pack(side=tk.LEFT, padx=5)
        ttk.Button(date_frame, text="Find Workbook", command=self.find_workbook_by_date).pack(side=tk.LEFT, padx=10)
        
        file_frame.columnconfigure(1, weight=1)
        
        # === INVOICE HEADER SECTION ===
        header_frame = ttk.LabelFrame(main_frame, text="Invoice Header", padding="10")
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Header fields (4 columns layout for compactness)
        # Row 1: Vendor Code, Challan No, Challan Date, Invoice No
        # Row 2: Invoice Date, Excise Amount, Sales Tax, OE/Spares
        
        self.header_entries = {}
        
        # Row 0
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
        
        # Row 1
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
        
        # Row 2 - OE/Spares Radio Buttons
        ttk.Label(header_frame, text="OE/Spares:", style="Header.TLabel").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        
        self.oe_spares_var = tk.StringVar(value="OE")  # Default to OE
        radio_frame = ttk.Frame(header_frame)
        radio_frame.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Radiobutton(radio_frame, text="O/E", variable=self.oe_spares_var, value="OE").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(radio_frame, text="Spare", variable=self.oe_spares_var, value="Spare").pack(side=tk.LEFT, padx=5)
        
        # === ACTION BUTTONS ===
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        buttons = [
            ("Load Preview", self.load_preview, "Primary.TButton"),
            ("Generate All", self.generate_all_spool, "Primary.TButton"),
            ("Save", self.save_changes, "Action.TButton"),
            ("Clear", self.clear_all, "Action.TButton"),
            ("View Output", self.view_output, "Action.TButton"),
            ("Close", self.root.quit, "Action.TButton"),
        ]
        
        for text, command, style in buttons:
            ttk.Button(button_frame, text=text, command=command, style=style).pack(side=tk.LEFT, padx=5)
        
        # === PREVIEW TABLE SECTION ===
        table_frame = ttk.LabelFrame(main_frame, text="Preview Table (Double-click to edit)", padding="5")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Navigation frame for multiple invoices
        nav_frame = ttk.Frame(table_frame)
        nav_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.prev_btn = ttk.Button(nav_frame, text="◀ Previous", command=self.prev_preview, state='disabled')
        self.prev_btn.pack(side=tk.LEFT, padx=5)
        
        self.preview_label_var = tk.StringVar(value="No invoices loaded")
        ttk.Label(nav_frame, textvariable=self.preview_label_var, font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=20)
        
        self.next_btn = ttk.Button(nav_frame, text="Next ▶", command=self.next_preview, state='disabled')
        self.next_btn.pack(side=tk.LEFT, padx=5)
        
        # Create Treeview with scrollbars
        tree_container = ttk.Frame(table_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        # Horizontal scrollbar
        h_scroll = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Vertical scrollbar
        v_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        columns = [col[0] for col in self.COLUMNS]
        self.tree = ttk.Treeview(tree_container, columns=columns, show='headings',
                                  xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        h_scroll.config(command=self.tree.xview)
        v_scroll.config(command=self.tree.yview)
        
        # Configure columns
        for col_id, col_name, col_width in self.COLUMNS:
            self.tree.heading(col_id, text=col_name)
            self.tree.column(col_id, width=col_width, minwidth=50)
        
        # Bind double-click for editing
        self.tree.bind('<Double-1>', self.on_cell_double_click)
        
        # === FOOTER / STATUS BAR ===
        # Distinct footer area
        footer_frame = ttk.Frame(self.root, style="Status.TFrame", padding=(10, 5))
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Copyright (Left)
        ttk.Label(footer_frame, text="© Nagarkot Forwarders Pvt Ltd", style="Copyright.TLabel").pack(side=tk.LEFT)
        
        # Status Message (Center/Right)
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(footer_frame, textvariable=self.status_var, style="Info.TLabel")
        self.status_bar.pack(side=tk.RIGHT, padx=10)
    
    def show_status(self, message, style='normal'):
        """Display status message with appropriate styling."""
        self.status_var.set(message)
        if style == 'success':
            self.status_bar.configure(style='Success.TLabel')
        elif style == 'error':
            self.status_bar.configure(style='Error.TLabel')
        elif style == 'info':
            self.status_bar.configure(style='Info.TLabel')
        else:
            self.status_bar.configure(style='Info.TLabel')
        self.root.update()
    
    def browse_invoices(self):
        """Browse for one or more invoice PDF files."""
        # Use last browsed directory or user's home
        if self.invoice_path and os.path.exists(os.path.dirname(self.invoice_path)):
            initial_dir = os.path.dirname(self.invoice_path)
        else:
            initial_dir = os.path.expanduser("~")
            
        filepaths = filedialog.askopenfilenames(
            title="Select Invoice PDF(s) - Hold Ctrl to select multiple",
            initialdir=initial_dir,
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filepaths:
            self.invoice_paths = list(filepaths)
            self.invoice_path = filepaths[0]
            if len(filepaths) == 1:
                self.invoice_path_var.set(filepaths[0])
                self.status_var.set(f"Invoice selected: {os.path.basename(filepaths[0])}")
            else:
                self.invoice_path_var.set(f"{len(filepaths)} invoices selected")
                self.status_var.set(f"{len(filepaths)} invoice(s) selected")
    
    def browse_excel(self):
        """Browse for Excel/CSV file."""
        # Use last browsed directory or user's home
        if self.excel_path and os.path.exists(os.path.dirname(self.excel_path)):
            initial_dir = os.path.dirname(self.excel_path)
        else:
            initial_dir = os.path.expanduser("~")
            
        filepath = filedialog.askopenfilename(
            title="Select Excel/CSV File",
            initialdir=initial_dir,
            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if filepath:
            self.excel_path = filepath
            self.excel_path_var.set(filepath)
            self.status_var.set(f"Excel selected: {os.path.basename(filepath)}")
    
    def find_workbook_by_date(self):
        """Find and select workbook based on selected date in the same directory as current Excel file."""
        try:
            self.selected_date = datetime.strptime(self.date_var.get().strip(), '%d-%m-%Y')
        except ValueError:
            messagebox.showerror("Error", "Invalid date format. Use DD-MM-YYYY")
            return
        
        # Use directory of currently selected Excel file, or ask user to browse first
        if self.excel_path and os.path.exists(os.path.dirname(self.excel_path)):
            search_dir = os.path.dirname(self.excel_path)
        else:
            messagebox.showinfo("Info", "Please browse and select an Excel/CSV file first.\n\n"
                               "The 'Find Workbook' will then search in that folder for the date.")
            return
        
        if not os.path.exists(search_dir):
            messagebox.showerror("Error", f"Directory not found: {search_dir}")
            return
        
        # Date format patterns to search for in filenames
        date_patterns = [
            self.selected_date.strftime('%d-%m-%Y'),  # 29-01-2026
            self.selected_date.strftime('%d-%m-%y'),  # 29-01-26
            self.selected_date.strftime('%d/%m/%Y'),  # 29/01/2026
            self.selected_date.strftime('%Y-%m-%d'),  # 2026-01-29
        ]
        
        found_file = None
        
        # First, look for CSV files with date in filename
        for filename in os.listdir(search_dir):
            filepath = os.path.join(search_dir, filename)
            if os.path.isfile(filepath):
                filename_upper = filename.upper()
                for pattern in date_patterns:
                    if pattern.upper() in filename_upper:
                        found_file = filepath
                        break
            if found_file:
                break
        
        # If no CSV found, look for Excel files with month/year in filename
        if not found_file:
            month_patterns = [
                self.selected_date.strftime('%b-%Y').upper(),  # JAN-2026
                self.selected_date.strftime('%B-%Y').upper(),  # JANUARY-2026
                self.selected_date.strftime('%m-%Y'),  # 01-2026
            ]
            for filename in os.listdir(search_dir):
                filepath = os.path.join(search_dir, filename)
                if os.path.isfile(filepath) and filename.lower().endswith(('.xlsx', '.xls')):
                    filename_upper = filename.upper()
                    for pattern in month_patterns:
                        if pattern in filename_upper:
                            found_file = filepath
                            break
                if found_file:
                    break
        
        if found_file:
            self.excel_path = found_file
            self.excel_path_var.set(found_file)
            # Non-intrusive positive feedback via status bar
            self.show_status(f"✓ Workbook found: {os.path.basename(found_file)} for {self.selected_date.strftime('%d-%m-%Y')}", 'success')
        else:
            # Error popup for not found
            messagebox.showerror("Workbook Not Found", 
                f"No workbook found for date {self.selected_date.strftime('%d-%m-%Y')}\n\n"
                f"Search location: {search_dir}\n\n"
                "Please verify:\n"
                "• The date is correct\n"
                "• The workbook exists in the correct folder\n"
                "• OE/Spare selection matches your workbook location")
            self.show_status(f"✗ Workbook not found for {self.selected_date.strftime('%d-%m-%Y')}", 'error')
    
    def normalize_item_code(self, code):
        """Normalize item code by removing dashes and spaces."""
        if not code:
            return ''
        return str(code).replace('-', '').replace(' ', '').upper()
    
    def extract_invoice_data(self):
        """Extract data from invoice PDF."""
        if not self.invoice_path:
            return False
        
        self.invoice_data = {}
        self.invoice_line_items = {}
        self.invoice_validation_info = {}  # Store validation-related info
        
        try:
            with pdfplumber.open(self.invoice_path) as pdf:
                full_text = ""
                has_images = False
                
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        full_text += page_text + "\n"
                    
                    # Check for images (QR codes are images)
                    if page.images:
                        has_images = True
                
                # Store info for validation
                self.invoice_validation_info['full_text'] = full_text
                self.invoice_validation_info['has_images'] = has_images
                self.invoice_validation_info['pdf_path'] = self.invoice_path
                
                # Invoice Number
                inv_match = re.search(r'Invoice\s+Number\s*:\s*([A-Z0-9/\-]+)', full_text, re.IGNORECASE)
                if inv_match:
                    self.invoice_data['invoice_no'] = inv_match.group(1).strip()
                
                # Invoice Date
                date_match = re.search(r'Invoice\s+Date\s*:\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})', full_text, re.IGNORECASE)
                if date_match:
                    self.invoice_data['invoice_date'] = date_match.group(1).strip()
                
                # PO Number
                po_match = re.search(r'Cust\s+PO\s+No\.?\s*:\s*(\d+)', full_text, re.IGNORECASE)
                if po_match:
                    self.invoice_data['po_number'] = po_match.group(1).strip()
                
                # Vendor Code
                ref_match = re.search(r'Reference\s+No\.?\s*:\s*([A-Z]\d{3,})', full_text, re.IGNORECASE)
                if ref_match:
                    self.invoice_data['vendor_code'] = ref_match.group(1).strip()
                
                # GST Number
                gst_match = re.search(r'GSTIN\s+Number\s*:\s*(\d{2}[A-Z]{5}\d{4}[A-Z]\d[A-Z\d]{2})', full_text, re.IGNORECASE)
                if gst_match:
                    self.invoice_data['gst_no'] = gst_match.group(1).strip()
                
                # IRN Number - 64 character hex string
                irn_match = re.search(r'IRN\s*(?:NO)?[:\s]*([a-f0-9]{64})', full_text, re.IGNORECASE)
                if irn_match:
                    self.invoice_data['irn_number'] = irn_match.group(1).strip()
                
                # Total Invoice Value
                total_match = re.search(r'Invoice\s+Amount\s*\(INR\)\s*([\d,]+\.?\d*)', full_text, re.IGNORECASE)
                if total_match:
                    self.invoice_data['total_value'] = total_match.group(1).replace(',', '').strip()
                
                # Tax totals
                tax_totals = re.search(
                    r'^([\d,]{6,}\.[0-9]{2})\s+([\d,]+\.[0-9]{2})\s+([\d,]+\.[0-9]{2})\s+([\d,]+\.[0-9]{2})\s+[\d,]+\.[0-9]{2}\s*$',
                    full_text, re.MULTILINE
                )
                if tax_totals:
                    self.invoice_data['cgst_amt'] = tax_totals.group(2).replace(',', '')
                    self.invoice_data['sgst_amt'] = tax_totals.group(3).replace(',', '')
                    igst_val = tax_totals.group(4).replace(',', '')
                    self.invoice_data['igst_amt'] = '' if float(igst_val) == 0 else igst_val
                
                # Line items
                line_pattern = re.compile(
                    r'(\d)\s+(\d{6}-\d{5})\s+(\d{8})\s+(\d+)\.00\s+Nos\s+([\d,]+\.\d{3})',
                    re.IGNORECASE
                )
                # Item code pattern: 6 alphanumeric chars + optional dash + 5 alphanumeric chars + optional suffix like -999
                # e.g., 56110M-58UA1, 5A120M58U01, 55101M-58U10, 55401M-58U00-999
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
                    if i + 1 < len(line_matches):
                        search_end = line_matches[i + 1].start()
                    else:
                        search_end = len(full_text)
                    
                    search_text = full_text[search_start:search_end]
                    item_code_match = item_code_pattern.search(search_text)
                    
                    if item_code_match:
                        item_code_raw = item_code_match.group(1)
                        item_code_norm = self.normalize_item_code(item_code_raw)
                        
                        self.invoice_line_items[item_code_norm] = {
                            'sno': sno,
                            'material_code': material_code,
                            'item_code': item_code_raw,
                            'hsn_code': hsn_code,
                            'qty': qty,
                            'rate': rate,
                        }
                
                self.invoice_data['eway_bill'] = '0'
            
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read invoice: {e}")
            return False
    
    def validate_invoice_integrity(self):
        """Validate invoice for required elements: IRN, Original copy, Digital signature, GSTIN.
        
        Returns:
            tuple: (is_valid, list of error messages)
        """
        errors = []
        full_text = self.invoice_validation_info.get('full_text', '')
        pdf_path = self.invoice_validation_info.get('pdf_path', '')
        
        # 1. IRN Number validation - must be 64 character alphanumeric hex string
        # Note: QR code validation removed - QR encodes the IRN, so if IRN exists, invoice is e-invoice compliant
        irn = self.invoice_data.get('irn_number', '')
        if not irn:
            errors.append("IRN Number not found in invoice")
        elif len(irn) != 64:
            errors.append(f"IRN Number invalid length: {len(irn)} chars (expected 64)")
        elif not re.match(r'^[a-f0-9]{64}$', irn, re.IGNORECASE):
            errors.append("IRN Number contains invalid characters (must be alphanumeric hex)")
        
        # 2. Original for Recipient - must be original copy, not duplicate
        original_patterns = [
            r'Original\s+for\s*\n?\s*Recipient',
            r'Original\s+for\s+Recipient',
            r'Tax\s+Invoice\s+Original',
        ]
        is_original = False
        for pattern in original_patterns:
            if re.search(pattern, full_text, re.IGNORECASE):
                is_original = True
                break
        if not is_original:
            errors.append("Invoice is not 'Original for Recipient' copy")
        
        # 3. Digital Signature validation
        # The signature text is often rendered as an overlay/XObject, not extractable by pdfplumber
        # Use PyMuPDF (fitz) for better extraction, or check PDF annotations for signature objects
        has_digital_signature = self._check_digital_signature(pdf_path, full_text)
        
        if not has_digital_signature:
            errors.append("Digital Signature not found (must have 'Digitally signed by...' with signer name)")
        
        # 5. GSTIN validation - format: 2 digits + 5 uppercase + 4 digits + 1 uppercase + 1 digit + 1 alphanumeric + 1 alphanumeric
        # Standard format: 22AAAAA0000A1Z5 (15 characters)
        gst_no = self.invoice_data.get('gst_no', '')
        if not gst_no:
            errors.append("GSTIN Number not found in invoice")
        else:
            # GSTIN format: SS PPPPP NNNN P E Z C
            # SS = State code (01-37), PPPPP = PAN first 5 (letters), NNNN = PAN next 4 (digits)
            # P = PAN last (letter), E = Entity (1-9), Z = default 'Z', C = checksum
            gstin_pattern = r'^[0-3][0-9][A-Z]{5}[0-9]{4}[A-Z][1-9A-Z][Z][0-9A-Z]$'
            if not re.match(gstin_pattern, gst_no, re.IGNORECASE):
                errors.append(f"GSTIN Number format invalid: {gst_no}")
            else:
                # Validate state code (01 to 37)
                state_code = int(gst_no[:2])
                if state_code < 1 or state_code > 37:
                    errors.append(f"GSTIN state code invalid: {state_code} (must be 01-37)")
        
        return len(errors) == 0, errors
    
    def _check_digital_signature(self, pdf_path, pdfplumber_text):
        """Check for digital signature using multiple methods.
        
        Args:
            pdf_path: Path to the PDF file
            pdfplumber_text: Text already extracted by pdfplumber
            
        Returns:
            bool: True if digital signature is found
        """
        # Method 1: Check text extracted by pdfplumber (may not work for signature overlays)
        digital_sign_patterns = [
            r'Digitally\s+signed\s+by\s+[A-Z\s]+',
            r'Digitally\s+signed\s+by.*Date:\d{4}\.\d{2}\.\d{2}',
            r'Digital\s+Signature.*Date:',
        ]
        for pattern in digital_sign_patterns:
            if re.search(pattern, pdfplumber_text, re.IGNORECASE):
                return True
        
        # Method 2: Use PyMuPDF (fitz) for better text extraction
        if HAS_PYMUPDF and pdf_path and os.path.exists(pdf_path):
            try:
                doc = fitz.open(pdf_path)
                for page in doc:
                    # Check for signature widgets (form fields)
                    widgets = list(page.widgets())
                    for widget in widgets:
                        if widget.field_type_string == 'Signature':
                            doc.close()
                            return True
                    
                    # Get text using PyMuPDF which extracts overlay text
                    text_dict = page.get_text("dict")
                    all_text = []
                    for block in text_dict.get("blocks", []):
                        if "lines" in block:
                            for line in block["lines"]:
                                for span in line["spans"]:
                                    all_text.append(span.get("text", ""))
                    
                    page_text = " ".join(all_text)
                    for pattern in digital_sign_patterns:
                        if re.search(pattern, page_text, re.IGNORECASE):
                            doc.close()
                            return True
                doc.close()
            except Exception:
                pass  # Fall through to Method 3
        
        # Method 3: Check pdfplumber annotations for signature objects
        if pdf_path and os.path.exists(pdf_path):
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    for page in pdf.pages:
                        annots = page.annots
                        if annots:
                            for annot in annots:
                                # Check if annotation is a signature widget
                                annot_data = annot.get('data', {})
                                if annot_data.get('FT') == '/Sig':
                                    # Found a signature field, check if it has signer name
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
    def load_excel_data(self):
        """Load and filter Excel/CSV data based on OE/Spare selection."""
        if not self.excel_path:
            return None
        
        try:
            is_spare = self.oe_spares_var.get() == "Spare"
            
            if self.excel_path.lower().endswith('.csv'):
                # Spare format has no header rows to skip, OE has 2 header rows
                if is_spare:
                    df = pd.read_csv(self.excel_path)
                else:
                    df = pd.read_csv(self.excel_path, skiprows=2)
            else:
                # Excel: pick worksheet by user-selected date
                xls = pd.ExcelFile(self.excel_path)
                sheet_names = xls.sheet_names
                
                # Use selected date or default to today
                search_date = self.selected_date if self.selected_date else datetime.now()

                def normalize(name):
                    return str(name).strip().upper()

                # Date tokens for the selected date
                date_tokens = {
                    search_date.strftime('%d-%m-%Y'),
                    search_date.strftime('%d/%m/%Y'),
                    search_date.strftime('%d-%b-%Y'),
                    search_date.strftime('%d-%b-%y'),
                    search_date.strftime('%d-%m-%y'),
                }

                candidates = []
                for name in sheet_names:
                    upper_name = normalize(name)
                    is_match = any(token.upper() in upper_name for token in date_tokens)
                    
                    if is_spare:
                        # For Spare/RPDC: include ALL sheets matching the date
                        # This covers both 'RPDC 20-01-2026' and '20-01-2026' sheets
                        if is_match:
                            candidates.append(name)
                    else:
                        # For OE: exclude RPDC sheets, only regular date sheets
                        if is_match and 'RPDC' not in upper_name:
                            candidates.append(name)
                
                print(f"DEBUG: Available sheets = {sheet_names}")
                print(f"DEBUG: Search date = {search_date.strftime('%d-%m-%Y')}")
                print(f"DEBUG: Is Spare = {is_spare}")
                print(f"DEBUG: Matching candidates = {candidates}")

                # Load all candidate sheets and concatenate
                dfs = []
                for sheet_name in candidates:
                    if is_spare:
                        sheet_df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
                    else:
                        sheet_df = pd.read_excel(self.excel_path, sheet_name=sheet_name, skiprows=2)
                    dfs.append(sheet_df)
                
                # Fallback if no candidates found
                if not dfs:
                    if is_spare:
                        for name in sheet_names:
                            if 'RPDC' in normalize(name):
                                dfs.append(pd.read_excel(self.excel_path, sheet_name=name))
                                break
                    if not dfs and sheet_names:
                        if is_spare:
                            dfs.append(pd.read_excel(self.excel_path, sheet_name=sheet_names[0]))
                        else:
                            dfs.append(pd.read_excel(self.excel_path, sheet_name=sheet_names[0], skiprows=2))
                
                if dfs:
                    df = pd.concat(dfs, ignore_index=True)
                else:
                    raise ValueError("No sheets found to load")
            
            df.columns = df.columns.str.strip().str.upper()
            return df
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel: {e}")
            return None
    
    def get_invoice_item(self, part_number):
        """Get invoice line item data for a part number."""
        if not part_number:
            return None
        
        part_norm = self.normalize_item_code(part_number)
        
        if part_norm in self.invoice_line_items:
            return self.invoice_line_items[part_norm]
        
        for key, value in self.invoice_line_items.items():
            if key in part_norm or part_norm in key:
                return value
        
        return None
    
    def load_preview(self):
        """Load data and populate preview table for all selected invoices."""
        if not self.invoice_path and not self.invoice_paths:
            messagebox.showwarning("Warning", "Please select Invoice PDF(s) first!")
            return
        if not self.excel_path:
            messagebox.showwarning("Warning", "Please select an Excel/CSV file first!")
            return
        
        # Use invoice_paths if available, otherwise single invoice_path
        invoices_to_load = self.invoice_paths if self.invoice_paths else [self.invoice_path]
        
        self.status_var.set(f"Loading {len(invoices_to_load)} invoice(s)...")
        self.root.update()
        
        # Load Excel data once
        excel_df = self.load_excel_data()
        if excel_df is None:
            return
        
        # Clear previous previews
        self.all_previews = []
        self.current_preview_index = 0
        
        # Track failures for better error messages
        validation_failures = []  # List of (filename, [errors])
        no_data_failures = []
        
        is_spare = self.oe_spares_var.get() == "Spare"
        
        # Get correct column names
        part_col = 'PART NUMBER'
        schedule_col = 'DI NUMBER' if is_spare else 'KANBAN NO'
        qty_col = 'SCHEDULED QUANTITY' if is_spare else 'QTY REQ'
        packing_col = 'PACKING STANDERD'
        batch_col = 'LATEST BATCH CODE' if is_spare else None
        
        for inv_idx, inv_path in enumerate(invoices_to_load):
            self.status_var.set(f"Loading invoice {inv_idx + 1}/{len(invoices_to_load)}...")
            self.root.update()
            
            # Set current invoice path and extract data
            self.invoice_path = inv_path
            if not self.extract_invoice_data():
                continue
            
            # Validate invoice integrity (IRN, QR Code, Original, Signature, GSTIN)
            is_valid, validation_errors = self.validate_invoice_integrity()
            if not is_valid:
                inv_filename = os.path.basename(inv_path)
                validation_failures.append((inv_filename, validation_errors))
                continue
            
            # Get invoice number for filtering
            inv_num = self.invoice_data.get('invoice_no', '').strip()
            if not inv_num:
                continue
            
            # Filter Excel by invoice number
            if 'INVOICE NO' not in excel_df.columns:
                continue
            
            matching = excel_df[excel_df['INVOICE NO'].astype(str).str.strip() == inv_num]
            
            # Filter by KANBAN NO or DI NUMBER
            if is_spare:
                if 'DI NUMBER' in matching.columns:
                    matching = matching[matching['DI NUMBER'].notna()]
            else:
                if 'KANBAN NO' in matching.columns:
                    matching = matching[matching['KANBAN NO'].notna()]
            
            # Filter by part numbers in invoice
            valid_rows = []
            for _, row in matching.iterrows():
                part_number = str(row.get(part_col, '')) if pd.notna(row.get(part_col)) else ''
                if self.get_invoice_item(part_number):
                    valid_rows.append(row)
            
            if not valid_rows:
                no_data_failures.append(f"{os.path.basename(inv_path)} ({inv_num})")
                continue
            
            # Validate quantities: compare invoice qty vs Nagare qty
            # First, aggregate Excel quantities by Part Number (multiple rows may exist for same part)
            excel_qty_by_part = {}
            for row in valid_rows:
                part_number = str(row.get(part_col, '')) if pd.notna(row.get(part_col)) else ''
                if not part_number:
                    continue
                
                excel_qty_val = row.get(qty_col, 0)
                try:
                    excel_qty = int(float(excel_qty_val)) if pd.notna(excel_qty_val) and excel_qty_val else 0
                except (ValueError, TypeError):
                    excel_qty = 0
                
                # Sum quantities for same part number
                if part_number in excel_qty_by_part:
                    excel_qty_by_part[part_number] += excel_qty
                else:
                    excel_qty_by_part[part_number] = excel_qty
            
            # Now compare aggregated Excel qty with invoice qty
            qty_mismatches = []
            for part_number, total_excel_qty in excel_qty_by_part.items():
                invoice_item = self.get_invoice_item(part_number)
                if invoice_item:
                    invoice_qty_str = invoice_item.get('qty', '0')
                    try:
                        invoice_qty = int(float(invoice_qty_str)) if invoice_qty_str else 0
                    except (ValueError, TypeError):
                        invoice_qty = 0
                    
                    if invoice_qty != total_excel_qty:
                        qty_mismatches.append({
                            'part': part_number,
                            'invoice_qty': invoice_qty,
                            'excel_qty': total_excel_qty
                        })
            
            if qty_mismatches:
                # Build detailed error message
                inv_name = os.path.basename(inv_path)
                error_lines = [f"Invoice: {inv_name}", "", "Quantity Mismatches Found:", ""]
                for mismatch in qty_mismatches:
                    error_lines.append(f"Part: {mismatch['part']}")
                    error_lines.append(f"  Invoice Qty: {mismatch['invoice_qty']}")
                    error_lines.append(f"  Nagare Qty: {mismatch['excel_qty']}")
                    error_lines.append("")
                
                messagebox.showerror("Quantity Mismatch", "\n".join(error_lines))
                continue  # Skip this invoice
            
            # Build preview data for this invoice
            preview_data = []
            for idx, row in enumerate(valid_rows):
                part_number = str(row.get(part_col, ''))
                invoice_item = self.get_invoice_item(part_number)
                
                unload_no = f"{datetime.now().strftime('%Y%m%d')}{idx+1:02d}"
                schedule_no = str(row.get(schedule_col, '')) if pd.notna(row.get(schedule_col)) else ''
                qty_val = row.get(qty_col, 0)
                qty = str(int(float(qty_val))) if pd.notna(qty_val) and qty_val else ''
                pack_val = row.get(packing_col, 0)
                bin_qty = str(int(float(pack_val))) if pd.notna(pack_val) and pack_val else ''
                batch_no = ''
                if is_spare and batch_col and batch_col in row.index:
                    batch_val = row.get(batch_col, '')
                    batch_no = str(batch_val) if pd.notna(batch_val) else ''
                
                row_data = {
                    'unload_no': unload_no,
                    'schedule_no': schedule_no,
                    'item_code': part_number,
                    'qty': qty,
                    'po_number': self.invoice_data.get('po_number', ''),
                    'f57_2no': '',
                    'bin_qty': bin_qty,
                    'remarks': '',
                    'batch_no': batch_no,
                    'location': '',
                    'gst_no': self.invoice_data.get('gst_no', ''),
                    'hsn_code': invoice_item.get('hsn_code', '') if invoice_item else '',
                    'cgst_amt': self.invoice_data.get('cgst_amt', ''),
                    'sgst_amt': self.invoice_data.get('sgst_amt', ''),
                    'igst_amt': self.invoice_data.get('igst_amt', ''),
                    'eway_bill': self.invoice_data.get('eway_bill', '0'),
                    'basic_price': invoice_item.get('rate', '') if invoice_item else '',
                    'total_value': self.invoice_data.get('total_value', ''),
                    'tool_amort': '0',
                }
                preview_data.append(row_data)
            
            # Store header data
            header_data = {
                'invoice_no': self.invoice_data.get('invoice_no', ''),
                'invoice_date': self.invoice_data.get('invoice_date', ''),
                'vendor_code': self.invoice_data.get('vendor_code', ''),
                'po_number': self.invoice_data.get('po_number', ''),
                'gst_no': self.invoice_data.get('gst_no', ''),
                'challan_no': self.invoice_data.get('invoice_no', ''),
                'challan_date': self.invoice_data.get('invoice_date', ''),
            }
            
            # Add to all_previews
            self.all_previews.append({
                'invoice_path': inv_path,
                'invoice_data': dict(self.invoice_data),
                'invoice_line_items': dict(self.invoice_line_items),
                'header_data': header_data,
                'preview_data': preview_data,
            })
        
        # Show validation errors for all failed invoices together
        if validation_failures:
            error_parts = []
            for inv_filename, errors in validation_failures:
                error_parts.append(f"📄 {inv_filename}:\n   • " + "\n   • ".join(errors))
            
            full_error_msg = "Invoice Validation Failed\n\n" + "\n\n".join(error_parts[:5])
            if len(validation_failures) > 5:
                full_error_msg += f"\n\n... and {len(validation_failures) - 5} more invoice(s)"
            
            messagebox.showerror("Invoice Validation Failed", full_error_msg)
        
        if not self.all_previews:
            # Provide detailed feedback about why no invoices were loaded
            if validation_failures and not no_data_failures:
                # All invoices failed validation - already shown error above
                self.show_status("All invoices failed validation", "error")
                return
            elif no_data_failures and not validation_failures:
                # All invoices had no matching Excel data
                msg = f"No matching data found in Excel for:\n• " + "\n• ".join(no_data_failures[:5])
                if len(no_data_failures) > 5:
                    msg += f"\n... and {len(no_data_failures) - 5} more"
                self.show_status("No matching Excel data found", "error")
                messagebox.showwarning("Warning", msg)
                return
            elif validation_failures and no_data_failures:
                # Mix of failures - validation already shown, show data failures
                msg = f"Additionally, no matching Excel data for:\n• " + "\n• ".join(no_data_failures[:3])
                self.show_status("Invoice loading failed", "error")
                messagebox.showwarning("Warning", msg)
                return
            else:
                msg = "No matching data found for any invoice!"
                self.show_status("No matching data found", "error")
                messagebox.showwarning("Warning", msg)
                return
        
        # Show first preview
        self.current_preview_index = 0
        self.show_current_preview()
        self.update_nav_buttons()
        
        self.show_status(f"✓ Loaded {len(self.all_previews)} invoice(s) successfully. Quantities verified.", "success")
    
    def show_current_preview(self):
        """Display the current invoice preview in the table."""
        if not self.all_previews or self.current_preview_index >= len(self.all_previews):
            return
        
        current = self.all_previews[self.current_preview_index]
        
        # Restore invoice data
        self.invoice_data = current['invoice_data']
        self.invoice_line_items = current['invoice_line_items']
        self.preview_data = current['preview_data']
        
        # Update header fields
        for field, entry in self.header_entries.items():
            entry.delete(0, tk.END)
            if field in current['header_data']:
                entry.insert(0, current['header_data'][field])
            elif field in self.invoice_data:
                entry.insert(0, self.invoice_data[field])
        
        # Clear and populate table
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for row_data in self.preview_data:
            values = [row_data[col[0]] for col in self.COLUMNS]
            self.tree.insert('', 'end', values=values)
        
        # Update preview label
        inv_no = self.invoice_data.get('invoice_no', 'Unknown')
        self.preview_label_var.set(f"Invoice {self.current_preview_index + 1} of {len(self.all_previews)}: {inv_no}")
    
    def update_nav_buttons(self):
        """Enable/disable navigation buttons based on current position."""
        if len(self.all_previews) <= 1:
            self.prev_btn.config(state='disabled')
            self.next_btn.config(state='disabled')
        else:
            self.prev_btn.config(state='normal' if self.current_preview_index > 0 else 'disabled')
            self.next_btn.config(state='normal' if self.current_preview_index < len(self.all_previews) - 1 else 'disabled')
    
    def prev_preview(self):
        """Show previous invoice preview."""
        if self.current_preview_index > 0:
            # Save current edits before switching
            self.save_current_preview_edits()
            self.current_preview_index -= 1
            self.show_current_preview()
            self.update_nav_buttons()
    
    def next_preview(self):
        """Show next invoice preview."""
        if self.current_preview_index < len(self.all_previews) - 1:
            # Save current edits before switching
            self.save_current_preview_edits()
            self.current_preview_index += 1
            self.show_current_preview()
            self.update_nav_buttons()
    
    def save_current_preview_edits(self):
        """Save edits made to current preview back to all_previews."""
        if not self.all_previews or self.current_preview_index >= len(self.all_previews):
            return
        
        # Update preview_data from table
        updated_preview = []
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id)['values']
            row_data = {}
            for col_idx, (col_id, _, _) in enumerate(self.COLUMNS):
                row_data[col_id] = values[col_idx]
            updated_preview.append(row_data)
        
        self.all_previews[self.current_preview_index]['preview_data'] = updated_preview
        
        # Also save header field edits
        for field, entry in self.header_entries.items():
            value = entry.get().strip()
            self.all_previews[self.current_preview_index]['header_data'][field] = value
    
    def on_cell_double_click(self, event):
        """Handle double-click for inline editing."""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        
        column = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        
        if not row:
            return
        
        col_idx = int(column[1:]) - 1  # Column index (0-based)
        col_id = self.COLUMNS[col_idx][0]
        
        # Get cell bbox
        bbox = self.tree.bbox(row, column)
        if not bbox:
            return
        
        # Get current value
        item = self.tree.item(row)
        current_value = item['values'][col_idx]
        
        # Create entry widget for editing
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
            
            # Update preview data
            row_idx = self.tree.get_children().index(row)
            if row_idx < len(self.preview_data):
                self.preview_data[row_idx][col_id] = new_value
            
            entry.destroy()
        
        def cancel_edit(event=None):
            entry.destroy()
        
        entry.bind('<Return>', save_edit)
        entry.bind('<Escape>', cancel_edit)
        entry.bind('<FocusOut>', save_edit)
    
    def generate_spool_line(self, row_data):
        """Generate a single 390-character spool line."""
        line = [' '] * 390
        
        # Ensure all row_data values are strings
        row_data = {k: str(v) if v is not None else '' for k, v in row_data.items()}
        
        def set_field(start, end, value, align='left'):
            width = end - start
            value_str = str(value) if value else ''
            if align == 'left':
                formatted = value_str[:width].ljust(width)
            else:
                formatted = value_str[:width].rjust(width)
            for i, char in enumerate(formatted):
                if start + i < 390:
                    line[start + i] = char
        
        # Format invoice date - use full 4-digit year
        invoice_date_raw = self.header_entries['invoice_date'].get()
        challan_date_raw = self.header_entries['challan_date'].get() or invoice_date_raw
        
        # Format challan date (dd-Mon-yyyy format like 28-Jan-2026)
        challan_date = challan_date_raw
        if challan_date_raw:
            for fmt in ['%d-%b-%y', '%d-%b-%Y', '%d/%m/%Y', '%d-%m-%Y']:
                try:
                    dt = datetime.strptime(challan_date_raw, fmt)
                    challan_date = dt.strftime('%d-%b-%Y')  # Title case month
                    break
                except:
                    continue
        
        # Format invoice date (uppercase: 28-JAN-2026)
        invoice_date = invoice_date_raw
        if invoice_date_raw:
            for fmt in ['%d-%b-%y', '%d-%b-%Y', '%d/%m/%Y', '%d-%m-%Y']:
                try:
                    dt = datetime.strptime(invoice_date_raw, fmt)
                    invoice_date = dt.strftime('%d-%b-%Y').upper()  # Full year: 28-JAN-2026
                    break
                except:
                    continue
        
        # Field positions
        set_field(0, 4, self.header_entries['vendor_code'].get() or 'X539')
        set_field(4, 20, self.header_entries['challan_no'].get() or self.header_entries['invoice_no'].get())
        set_field(20, 31, challan_date)
        set_field(31, 47, self.header_entries['invoice_no'].get())
        set_field(47, 71, invoice_date)
        
        # OE/Spares prefix: O/E = "1", Spare = "S"
        is_spare = self.oe_spares_var.get() == "Spare"
        oe_prefix = '1' if not is_spare else 'S'
        set_field(82, 83, oe_prefix)
        set_field(83, 98, row_data['schedule_no'])
        set_field(98, 113, row_data['item_code'])
        set_field(113, 125, row_data['qty'])
        set_field(125, 138, self.header_entries['po_number'].get())  # PO Number
        bin_qty_val = str(row_data['bin_qty']).strip()
        bin_qty_field = f"    {bin_qty_val}" if bin_qty_val else ''
        set_field(138, 150, bin_qty_field)
        
        # Batch code for Spare (position 194-204)
        if is_spare and row_data.get('batch_no'):
            set_field(194, 204, str(row_data['batch_no']).strip())
        
        set_field(204, 219, row_data['gst_no'])
        set_field(219, 227, row_data['hsn_code'])
        
        cgst = str(row_data['cgst_amt']) if row_data['cgst_amt'] else '0'
        sgst = str(row_data['sgst_amt']) if row_data['sgst_amt'] else '0'
        cgst_portion = cgst.ljust(16) + sgst[:2]
        set_field(227, 245, cgst_portion)
        sgst_portion = sgst[2:] if len(sgst) > 2 else ''
        set_field(245, 265, sgst_portion)
        
        set_field(275, 276, str(row_data['eway_bill']) if row_data['eway_bill'] else '0')
        set_field(276, 290, row_data['igst_amt'])
        set_field(290, 354, self.invoice_data.get('irn_number', ''))
        set_field(354, 366, row_data['basic_price'])
        set_field(366, 390, row_data['total_value'])
        
        return ''.join(line)
    
    def validate_required_fields(self, preview_idx=None):
        """Validate that required header fields are filled.
        
        Args:
            preview_idx: If provided, validates a specific preview. Otherwise validates current header entries.
        
        Returns:
            tuple: (is_valid, missing_fields_message)
        """
        required_fields = {
            'vendor_code': 'Vendor Code',
            'challan_no': 'Challan No',
            'challan_date': 'Challan Date',
            'invoice_no': 'Invoice No',
            'invoice_date': 'Invoice Date',
            'po_number': 'PO Number',
        }
        
        missing = []
        
        if preview_idx is not None and preview_idx < len(self.all_previews):
            # Validate from stored preview data
            preview = self.all_previews[preview_idx]
            header_data = preview.get('header_data', {})
            invoice_data = preview.get('invoice_data', {})
            
            for field_id, field_name in required_fields.items():
                value = header_data.get(field_id, '') or invoice_data.get(field_id, '')
                if not value or not str(value).strip():
                    missing.append(field_name)
        else:
            # Validate from current header entries
            for field_id, field_name in required_fields.items():
                entry = self.header_entries.get(field_id)
                if entry:
                    value = entry.get().strip()
                    if not value:
                        missing.append(field_name)
                else:
                    missing.append(field_name)
        
        if missing:
            return False, f"Missing required fields:\n• " + "\n• ".join(missing)
        return True, ""
    
    def validate_preview_rows(self, preview_idx=None):
        """Validate that required row fields are filled based on OE/Spare type.
        
        Args:
            preview_idx: If provided, validates a specific preview's rows. Otherwise validates current preview_data.
        
        Returns:
            tuple: (is_valid, list of error messages)
        """
        is_spare = self.oe_spares_var.get() == "Spare"
        
        # Define required fields for each type
        # Common required fields for both OE and Spare - these are essential for spool generation
        common_required = [
            'unload_no',      # Used for tracking
            'schedule_no',    # Excel: KANBAN NO - critical for spool
            'item_code',      # Excel: PART NUMBER - critical for spool
            'qty',            # Invoice/Excel qty - critical for spool
            'po_number',      # Invoice: Cust PO No - used in spool
            'gst_no',         # Invoice: GSTIN - critical for spool
            'hsn_code',       # Invoice: HSN code per item - critical for spool
            'cgst_amt',       # Invoice: CGST amount - required for GST spool
            'sgst_amt',       # Invoice: SGST amount - required for GST spool
            'basic_price',    # Invoice: Rate per item - critical for spool
            'total_value',    # Invoice: Total invoice value - critical for spool
        ]
        
        # OE-specific required fields
        oe_required = common_required + ['bin_qty']
        
        # Spare-specific required fields
        spare_required = common_required + ['batch_no', 'bin_qty']
        
        required_fields = spare_required if is_spare else oe_required
        
        # Field display names
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
        
        # Get preview data
        if preview_idx is not None and preview_idx < len(self.all_previews):
            preview_data = self.all_previews[preview_idx].get('preview_data', [])
        else:
            preview_data = self.preview_data if hasattr(self, 'preview_data') else []
        
        # GSTIN format pattern: SS PPPPP NNNN P E Z C (15 chars)
        gstin_pattern = r'^[0-3][0-9][A-Z]{5}[0-9]{4}[A-Z][1-9A-Z][Z][0-9A-Z]$'
        
        for row_idx, row in enumerate(preview_data):
            row_errors = []
            for field_id in required_fields:
                value = row.get(field_id, '')
                if not value or not str(value).strip():
                    row_errors.append(field_names.get(field_id, field_id))
            
            # Additional format validation for GST No
            gst_no = row.get('gst_no', '')
            if gst_no and str(gst_no).strip():
                gst_no = str(gst_no).strip()
                if not re.match(gstin_pattern, gst_no, re.IGNORECASE):
                    row_errors.append(f"GST No format invalid ({gst_no})")
                else:
                    # Validate state code (01 to 37)
                    state_code = int(gst_no[:2])
                    if state_code < 1 or state_code > 37:
                        row_errors.append(f"GST No state code invalid ({state_code})")
            
            if row_errors:
                item_code = row.get('item_code', f'Row {row_idx + 1}')
                errors.append(f"Row '{item_code}': {', '.join(row_errors)}")
        
        return len(errors) == 0, errors
    
    def generate_all_spool(self):
        # """Generate spool files for all loaded invoice previews."""
        # First save any current edits to the current preview
        self.save_current_preview_edits()
        
        if not self.all_previews:
            messagebox.showwarning("Warning", "No previews loaded!\nSelect invoice(s) and click 'Load Preview' first.")
            return
        
        # Validate required fields for all previews
        validation_errors = []
        for idx, preview in enumerate(self.all_previews):
            is_valid, error_msg = self.validate_required_fields(idx)
            if not is_valid:
                inv_no = preview.get('invoice_data', {}).get('invoice_no', f'Invoice {idx+1}')
                validation_errors.append(f"{inv_no}:\n{error_msg}")
        
        if validation_errors:
            error_display = "\n\n".join(validation_errors[:5])
            if len(validation_errors) > 5:
                error_display += f"\n\n... and {len(validation_errors) - 5} more invoice(s) with errors"
            messagebox.showerror("Header Validation Error", f"Cannot generate spool files.\n\n{error_display}")
            return
        
        # Validate preview row fields for all previews
        row_validation_errors = []
        for idx, preview in enumerate(self.all_previews):
            is_valid, row_errors = self.validate_preview_rows(idx)
            if not is_valid:
                inv_no = preview.get('invoice_data', {}).get('invoice_no', f'Invoice {idx+1}')
                row_validation_errors.append(f"{inv_no}:\n• " + "\n• ".join(row_errors[:3]))
                if len(row_errors) > 3:
                    row_validation_errors[-1] += f"\n  ... and {len(row_errors) - 3} more row(s)"
        
        if row_validation_errors:
            error_display = "\n\n".join(row_validation_errors[:3])
            if len(row_validation_errors) > 3:
                error_display += f"\n\n... and {len(row_validation_errors) - 3} more invoice(s) with row errors"
            messagebox.showerror("Row Data Validation Error", f"Cannot generate spool files - missing required row data.\n\n{error_display}")
            return
        
        # Determine output location
        is_spare = self.oe_spares_var.get() == "Spare"
        if is_spare:
            default_output = os.path.join(self.output_dir, "Spare")
        else:
            default_output = os.path.join(self.output_dir, "Original")
        os.makedirs(default_output, exist_ok=True)
        
        # For single invoice, use Save As dialog; for multiple, use folder selection
        if len(self.all_previews) == 1:
            # Single invoice - Save As dialog
            preview = self.all_previews[0]
            inv_no = preview['invoice_data'].get('invoice_no', '').strip()
            
            if not inv_no:
                messagebox.showerror("Error", "Invoice number is required!")
                return
            
            default_filename = inv_no.split('/')[-1] if '/' in inv_no else inv_no
            
            output_path = filedialog.asksaveasfilename(
                title="Save Spool File",
                initialdir=default_output,
                initialfile=f"{default_filename}.txt",
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
            
            if not output_path:
                return
            
            # Generate spool lines
            lines = []
            for row_data in preview['preview_data']:
                # Set invoice data for line generation
                self.invoice_data = preview['invoice_data']
                for field, entry in self.header_entries.items():
                    entry.delete(0, tk.END)
                    if field in preview['header_data']:
                        entry.insert(0, preview['header_data'][field])
                    elif field in self.invoice_data:
                        entry.insert(0, self.invoice_data[field])
                
                line = self.generate_spool_line(row_data)
                lines.append(line)
            
            try:
                with open(output_path, 'w', encoding='utf-8') as f:
                    for line in lines:
                        f.write(line + '\n')
                
                messagebox.showinfo("Success", f"Spool file saved:\n{output_path}\n\n{len(lines)} line(s) written.")
                self.status_var.set(f"Saved: {os.path.basename(output_path)} ({len(lines)} lines)")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to write file: {e}")
        else:
            # Multiple invoices - folder selection
            output_folder = filedialog.askdirectory(
                title="Select Output Folder for Spool Files",
                initialdir=default_output
            )
            
            if not output_folder:
                return
            
            success_count = 0
            error_count = 0
            error_invoices = []
            
            self.status_var.set(f"Generating {len(self.all_previews)} spool files...")
            self.root.update()
            
            for idx, preview in enumerate(self.all_previews):
                try:
                    inv_data = preview['invoice_data']
                    inv_no = inv_data.get('invoice_no', '').strip()
                    
                    if not inv_no:
                        error_count += 1
                        error_invoices.append(f"Invoice {idx+1} (no invoice number)")
                        continue
                    
                    self.status_var.set(f"Generating {idx+1}/{len(self.all_previews)}: {inv_no}")
                    self.root.update()
                    
                    # Set data for spool line generation
                    self.invoice_data = preview['invoice_data']
                    
                    # Update header entries
                    for field, entry in self.header_entries.items():
                        entry.delete(0, tk.END)
                        if field in preview['header_data']:
                            entry.insert(0, preview['header_data'][field])
                        elif field in self.invoice_data:
                            entry.insert(0, self.invoice_data[field])
                    
                    # Generate spool lines
                    lines = []
                    for row_data in preview['preview_data']:
                        line = self.generate_spool_line(row_data)
                        lines.append(line)
                    
                    # Save file
                    default_filename = inv_no.split('/')[-1] if '/' in inv_no else inv_no
                    output_path = os.path.join(output_folder, f"{default_filename}.txt")
                    
                    with open(output_path, 'w', encoding='utf-8') as f:
                        for line in lines:
                            f.write(line + '\n')
                    
                    success_count += 1
                    
                except Exception as e:
                    error_count += 1
                    error_invoices.append(f"Invoice {idx+1} ({str(e)})")
            
            # Restore current preview
            self.show_current_preview()
            
            # Show summary
            summary = f"Generation Complete!\n\nSuccess: {success_count}\nErrors: {error_count}"
            if error_invoices:
                summary += f"\n\nFailed invoices:\n" + "\n".join(error_invoices[:10])
                if len(error_invoices) > 10:
                    summary += f"\n... and {len(error_invoices) - 10} more"
            
            messagebox.showinfo("Generation Complete", summary)
            self.status_var.set(f"Generated: {success_count} success, {error_count} errors")
    
    def save_changes(self):
        # \"\"\"Save current table data to all_previews.\"\"\"
        # First save current preview edits to all_previews
        self.save_current_preview_edits()
        
        # Update preview_data from table (for backward compatibility)
        for idx, item_id in enumerate(self.tree.get_children()):
            values = self.tree.item(item_id)['values']
            if idx < len(self.preview_data):
                for col_idx, (col_id, _, _) in enumerate(self.COLUMNS):
                    self.preview_data[idx][col_id] = values[col_idx]
        
        messagebox.showinfo("Saved", "Changes saved successfully!")
        self.status_var.set("Changes saved")
    
    def clear_all(self):
        """Clear all data."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for entry in self.header_entries.values():
            entry.delete(0, tk.END)
        
        self.invoice_path_var.set('')
        self.excel_path_var.set('')
        self.invoice_path = None
        self.invoice_paths = []  # Clear bulk list
        self.excel_path = None
        self.preview_data = []
        self.invoice_data = {}
        self.invoice_line_items = {}
        self.oe_spares_var.set("OE")  # Reset to default
        
        # Reset multi-preview navigation
        self.all_previews = []
        self.current_preview_index = 0
        self.preview_label_var.set("No invoices loaded")
        self.prev_btn.config(state='disabled')
        self.next_btn.config(state='disabled')
        
        self.status_var.set("Cleared - Ready for new data")
    
    def view_output(self):
        """Open output directory."""
        if os.path.exists(self.output_dir):
            os.startfile(self.output_dir)
        else:
            messagebox.showwarning("Warning", "Output directory does not exist!")


def main():
    root = tk.Tk()
    app = SpoolFileGeneratorV2(root)
    root.mainloop()


if __name__ == "__main__":
    main()
