import os
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog

import pandas as pd

from .utils import get_initial_dir
from .services.invoice_service import extract_invoice_data, get_invoice_item, validate_invoice_integrity
from .services.excel_service import load_excel_data
from .services.validation_service import validate_required_fields, validate_preview_rows
from .services.spool_service import generate_spool_line


class SpoolAppController:
    def __init__(self, root, view):
        self.root = root
        self.view = view

        # Directories
        if getattr(__import__('sys'), 'frozen', False):
            self.base_dir = os.path.dirname(__import__('sys').executable)
        else:
            self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_dir = os.path.join(self.base_dir, "SpoolOutput")
        os.makedirs(self.output_dir, exist_ok=True)

        # Data storage
        self.invoice_path = None
        self.invoice_paths = []
        self.excel_path = None
        self.invoice_data = {}
        self.invoice_line_items = {}
        self.preview_data = []
        self.selected_date = None

        # Multi-invoice preview storage
        self.all_previews = []
        self.current_preview_index = 0

    def get_today(self):
        return datetime.now().strftime('%d-%m-%Y')

    def browse_invoices(self):
        initial_dir = get_initial_dir(self.invoice_path)
        filepaths = filedialog.askopenfilenames(
            title="Select Invoice PDF(s) - Hold Ctrl to select multiple",
            initialdir=initial_dir,
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filepaths:
            self.invoice_paths = list(filepaths)
            self.invoice_path = filepaths[0]
            if len(filepaths) == 1:
                self.view.set_invoice_path_display(filepaths[0])
                self.view.set_status(f"Invoice selected: {os.path.basename(filepaths[0])}")
            else:
                self.view.set_invoice_path_display(f"{len(filepaths)} invoices selected")
                self.view.set_status(f"{len(filepaths)} invoice(s) selected")

    def browse_excel(self):
        initial_dir = get_initial_dir(self.excel_path)
        filepath = filedialog.askopenfilename(
            title="Select Excel/CSV File",
            initialdir=initial_dir,
            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if filepath:
            self.excel_path = filepath
            self.view.set_excel_path_display(filepath)
            self.view.set_status(f"Excel selected: {os.path.basename(filepath)}")

    def find_workbook_by_date(self):
        try:
            self.selected_date = datetime.strptime(self.view.get_dispatch_date(), '%d-%m-%Y')
        except ValueError:
            messagebox.showerror("Error", "Invalid date format. Use DD-MM-YYYY")
            return

        if self.excel_path and os.path.exists(os.path.dirname(self.excel_path)):
            search_dir = os.path.dirname(self.excel_path)
        else:
            messagebox.showinfo("Info", "Please browse and select an Excel/CSV file first.\n\n"
                               "The 'Find Workbook' will then search in that folder for the date.")
            return

        if not os.path.exists(search_dir):
            messagebox.showerror("Error", f"Directory not found: {search_dir}")
            return

        date_patterns = [
            self.selected_date.strftime('%d-%m-%Y'),
            self.selected_date.strftime('%d-%m-%y'),
            self.selected_date.strftime('%d/%m/%Y'),
            self.selected_date.strftime('%Y-%m-%d'),
        ]

        found_file = None
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

        if not found_file:
            month_patterns = [
                self.selected_date.strftime('%b-%Y').upper(),
                self.selected_date.strftime('%B-%Y').upper(),
                self.selected_date.strftime('%m-%Y'),
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
            self.view.set_excel_path_display(found_file)
            self.view.set_status(f"✓ Workbook found: {os.path.basename(found_file)} for {self.selected_date.strftime('%d-%m-%Y')}", 'success')
        else:
            messagebox.showerror("Workbook Not Found",
                f"No workbook found for date {self.selected_date.strftime('%d-%m-%Y')}\n\n"
                f"Search location: {search_dir}\n\n"
                "Please verify:\n"
                "• The date is correct\n"
                "• The workbook exists in the correct folder\n"
                "• OE/Spare selection matches your workbook location")
            self.view.set_status(f"✗ Workbook not found for {self.selected_date.strftime('%d-%m-%Y')}", 'error')

    def _extract_invoice_data(self):
        if not self.invoice_path:
            return False
        try:
            invoice_data, invoice_line_items, validation_info = extract_invoice_data(self.invoice_path)
            self.invoice_data = invoice_data
            self.invoice_line_items = invoice_line_items
            self.invoice_validation_info = validation_info
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read invoice: {e}")
            return False

    def _load_excel_data(self):
        if not self.excel_path:
            return None
        try:
            is_spare = self.view.get_oe_spares() == "Spare"
            return load_excel_data(self.excel_path, is_spare, self.selected_date)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel: {e}")
            return None

    def load_preview(self):
        if not self.invoice_path and not self.invoice_paths:
            messagebox.showwarning("Warning", "Please select Invoice PDF(s) first!")
            return
        if not self.excel_path:
            messagebox.showwarning("Warning", "Please select an Excel/CSV file first!")
            return

        invoices_to_load = self.invoice_paths if self.invoice_paths else [self.invoice_path]

        self.view.set_status(f"Loading {len(invoices_to_load)} invoice(s)...")
        self.root.update()

        excel_df = self._load_excel_data()
        if excel_df is None:
            return

        self.all_previews = []
        self.current_preview_index = 0

        validation_failures = []
        no_data_failures = []

        is_spare = self.view.get_oe_spares() == "Spare"

        part_col = 'PART NUMBER'
        schedule_col = 'DI NUMBER' if is_spare else 'KANBAN NO'
        qty_col = 'SCHEDULED QUANTITY' if is_spare else 'QTY REQ'
        packing_col = 'PACKING STANDERD'
        batch_col = 'LATEST BATCH CODE' if is_spare else None

        for inv_idx, inv_path in enumerate(invoices_to_load):
            self.view.set_status(f"Loading invoice {inv_idx + 1}/{len(invoices_to_load)}...")
            self.root.update()

            self.invoice_path = inv_path
            if not self._extract_invoice_data():
                continue

            is_valid, validation_errors = validate_invoice_integrity(self.invoice_data, self.invoice_validation_info)
            if not is_valid:
                inv_filename = os.path.basename(inv_path)
                validation_failures.append((inv_filename, validation_errors))
                continue

            inv_num = self.invoice_data.get('invoice_no', '').strip()
            if not inv_num:
                continue

            if 'INVOICE NO' not in excel_df.columns:
                continue

            matching = excel_df[excel_df['INVOICE NO'].astype(str).str.strip() == inv_num]

            if is_spare:
                if 'DI NUMBER' in matching.columns:
                    matching = matching[matching['DI NUMBER'].notna()]
            else:
                if 'KANBAN NO' in matching.columns:
                    matching = matching[matching['KANBAN NO'].notna()]

            valid_rows = []
            for _, row in matching.iterrows():
                part_number = str(row.get(part_col, '')) if pd.notna(row.get(part_col)) else ''
                if get_invoice_item(part_number, self.invoice_line_items):
                    valid_rows.append(row)

            if not valid_rows:
                no_data_failures.append(f"{os.path.basename(inv_path)} ({inv_num})")
                continue

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

                if part_number in excel_qty_by_part:
                    excel_qty_by_part[part_number] += excel_qty
                else:
                    excel_qty_by_part[part_number] = excel_qty

            qty_mismatches = []
            for part_number, total_excel_qty in excel_qty_by_part.items():
                invoice_item = get_invoice_item(part_number, self.invoice_line_items)
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
                inv_name = os.path.basename(inv_path)
                error_lines = [f"Invoice: {inv_name}", "", "Quantity Mismatches Found:", ""]
                for mismatch in qty_mismatches:
                    error_lines.append(f"Part: {mismatch['part']}")
                    error_lines.append(f"  Invoice Qty: {mismatch['invoice_qty']}")
                    error_lines.append(f"  Nagare Qty: {mismatch['excel_qty']}")
                    error_lines.append("")

                messagebox.showerror("Quantity Mismatch", "\n".join(error_lines))
                continue

            preview_data = []
            for idx, row in enumerate(valid_rows):
                part_number = str(row.get(part_col, ''))
                invoice_item = get_invoice_item(part_number, self.invoice_line_items)

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

            header_data = {
                'invoice_no': self.invoice_data.get('invoice_no', ''),
                'invoice_date': self.invoice_data.get('invoice_date', ''),
                'vendor_code': self.invoice_data.get('vendor_code', ''),
                'po_number': self.invoice_data.get('po_number', ''),
                'gst_no': self.invoice_data.get('gst_no', ''),
                'challan_no': self.invoice_data.get('invoice_no', ''),
                'challan_date': self.invoice_data.get('invoice_date', ''),
            }

            self.all_previews.append({
                'invoice_path': inv_path,
                'invoice_data': dict(self.invoice_data),
                'invoice_line_items': dict(self.invoice_line_items),
                'header_data': header_data,
                'preview_data': preview_data,
            })

        if validation_failures:
            error_parts = []
            for inv_filename, errors in validation_failures:
                error_parts.append(f"📄 {inv_filename}:\n   • " + "\n   • ".join(errors))

            full_error_msg = "Invoice Validation Failed\n\n" + "\n\n".join(error_parts[:5])
            if len(validation_failures) > 5:
                full_error_msg += f"\n\n... and {len(validation_failures) - 5} more invoice(s)"

            messagebox.showerror("Invoice Validation Failed", full_error_msg)

        if not self.all_previews:
            if validation_failures and not no_data_failures:
                self.view.set_status("All invoices failed validation", "error")
                return
            elif no_data_failures and not validation_failures:
                msg = f"No matching data found in Excel for:\n• " + "\n• ".join(no_data_failures[:5])
                if len(no_data_failures) > 5:
                    msg += f"\n... and {len(no_data_failures) - 5} more"
                self.view.set_status("No matching Excel data found", "error")
                messagebox.showwarning("Warning", msg)
                return
            elif validation_failures and no_data_failures:
                msg = f"Additionally, no matching Excel data for:\n• " + "\n• ".join(no_data_failures[:3])
                self.view.set_status("Invoice loading failed", "error")
                messagebox.showwarning("Warning", msg)
                return
            else:
                msg = "No matching data found for any invoice!"
                self.view.set_status("No matching data found", "error")
                messagebox.showwarning("Warning", msg)
                return

        self.current_preview_index = 0
        self.show_current_preview()
        self.update_nav_buttons()

        self.view.set_status(f"✓ Loaded {len(self.all_previews)} invoice(s) successfully. Quantities verified.", "success")

    def show_current_preview(self):
        if not self.all_previews or self.current_preview_index >= len(self.all_previews):
            return

        current = self.all_previews[self.current_preview_index]

        self.invoice_data = current['invoice_data']
        self.invoice_line_items = current['invoice_line_items']
        self.preview_data = current['preview_data']

        header_data = current['header_data']
        combined = dict(header_data)
        for key in self.invoice_data:
            combined.setdefault(key, self.invoice_data[key])
        self.view.set_header_values(combined)

        self.view.set_table_rows(self.preview_data)

        inv_no = self.invoice_data.get('invoice_no', 'Unknown')
        self.view.set_preview_label(f"Invoice {self.current_preview_index + 1} of {len(self.all_previews)}: {inv_no}")

    def update_nav_buttons(self):
        if len(self.all_previews) <= 1:
            self.view.set_nav_state(False, False)
        else:
            self.view.set_nav_state(self.current_preview_index > 0,
                                    self.current_preview_index < len(self.all_previews) - 1)

    def prev_preview(self):
        if self.current_preview_index > 0:
            self.save_current_preview_edits()
            self.current_preview_index -= 1
            self.show_current_preview()
            self.update_nav_buttons()

    def next_preview(self):
        if self.current_preview_index < len(self.all_previews) - 1:
            self.save_current_preview_edits()
            self.current_preview_index += 1
            self.show_current_preview()
            self.update_nav_buttons()

    def save_current_preview_edits(self):
        if not self.all_previews or self.current_preview_index >= len(self.all_previews):
            return

        updated_preview = self.view.get_table_rows()
        self.all_previews[self.current_preview_index]['preview_data'] = updated_preview

        header_values = self.view.get_header_values()
        self.all_previews[self.current_preview_index]['header_data'].update(header_values)

    def generate_spool_line(self, row_data):
        header_values = self.view.get_header_values()
        is_spare = self.view.get_oe_spares() == "Spare"
        return generate_spool_line(row_data, header_values, self.invoice_data, is_spare)

    def _validate_headers(self):
        header_values = self.view.get_header_values()
        return validate_required_fields(header_values, self.invoice_data)

    def _validate_rows(self, preview_data):
        is_spare = self.view.get_oe_spares() == "Spare"
        return validate_preview_rows(preview_data, is_spare)

    def generate_all_spool(self):
        self.save_current_preview_edits()

        if not self.all_previews:
            messagebox.showwarning("Warning", "No previews loaded!\nSelect invoice(s) and click 'Load Preview' first.")
            return

        validation_errors = []
        for idx, preview in enumerate(self.all_previews):
            is_valid, error_msg = self._validate_headers()
            if not is_valid:
                inv_no = preview.get('invoice_data', {}).get('invoice_no', f'Invoice {idx+1}')
                validation_errors.append(f"{inv_no}:\n{error_msg}")

        if validation_errors:
            error_display = "\n\n".join(validation_errors[:5])
            if len(validation_errors) > 5:
                error_display += f"\n\n... and {len(validation_errors) - 5} more invoice(s) with errors"
            messagebox.showerror("Header Validation Error", f"Cannot generate spool files.\n\n{error_display}")
            return

        row_validation_errors = []
        for idx, preview in enumerate(self.all_previews):
            is_valid, row_errors = self._validate_rows(preview['preview_data'])
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

        is_spare = self.view.get_oe_spares() == "Spare"
        default_output = os.path.join(self.output_dir, "Spare" if is_spare else "Original")
        os.makedirs(default_output, exist_ok=True)

        if len(self.all_previews) == 1:
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

            lines = []
            for row_data in preview['preview_data']:
                self.invoice_data = preview['invoice_data']
                header_data = preview['header_data']
                self.view.set_header_values(header_data)

                line = self.generate_spool_line(row_data)
                lines.append(line)

            try:
                with open(output_path, 'w', encoding='utf-8') as f:
                    for line in lines:
                        f.write(line + '\n')

                messagebox.showinfo("Success", f"Spool file saved:\n{output_path}\n\n{len(lines)} line(s) written.")
                self.view.set_status(f"Saved: {os.path.basename(output_path)} ({len(lines)} lines)")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to write file: {e}")
        else:
            output_folder = filedialog.askdirectory(
                title="Select Output Folder for Spool Files",
                initialdir=default_output
            )

            if not output_folder:
                return

            success_count = 0
            error_count = 0
            error_invoices = []

            self.view.set_status(f"Generating {len(self.all_previews)} spool files...")
            self.root.update()

            for idx, preview in enumerate(self.all_previews):
                try:
                    inv_data = preview['invoice_data']
                    inv_no = inv_data.get('invoice_no', '').strip()

                    if not inv_no:
                        error_count += 1
                        error_invoices.append(f"Invoice {idx+1} (no invoice number)")
                        continue

                    self.view.set_status(f"Generating {idx+1}/{len(self.all_previews)}: {inv_no}")
                    self.root.update()

                    self.invoice_data = preview['invoice_data']
                    self.view.set_header_values(preview['header_data'])

                    lines = []
                    for row_data in preview['preview_data']:
                        line = self.generate_spool_line(row_data)
                        lines.append(line)

                    default_filename = inv_no.split('/')[-1] if '/' in inv_no else inv_no
                    output_path = os.path.join(output_folder, f"{default_filename}.txt")

                    with open(output_path, 'w', encoding='utf-8') as f:
                        for line in lines:
                            f.write(line + '\n')

                    success_count += 1

                except Exception as e:
                    error_count += 1
                    error_invoices.append(f"Invoice {idx+1} ({str(e)})")

            self.show_current_preview()

            summary = f"Generation Complete!\n\nSuccess: {success_count}\nErrors: {error_count}"
            if error_invoices:
                summary += f"\n\nFailed invoices:\n" + "\n".join(error_invoices[:10])
                if len(error_invoices) > 10:
                    summary += f"\n... and {len(error_invoices) - 10} more"

            messagebox.showinfo("Generation Complete", summary)
            self.view.set_status(f"Generated: {success_count} success, {error_count} errors")

    def save_changes(self):
        self.save_current_preview_edits()
        messagebox.showinfo("Saved", "Changes saved successfully!")
        self.view.set_status("Changes saved")

    def clear_all(self):
        self.view.clear_table()
        self.view.set_header_values({})
        self.view.set_invoice_path_display('')
        self.view.set_excel_path_display('')

        self.invoice_path = None
        self.invoice_paths = []
        self.excel_path = None
        self.preview_data = []
        self.invoice_data = {}
        self.invoice_line_items = {}
        self.all_previews = []
        self.current_preview_index = 0

        self.view.set_preview_label("No invoices loaded")
        self.view.set_nav_state(False, False)
        self.view.set_status("Cleared - Ready for new data")

    def view_output(self):
        if os.path.exists(self.output_dir):
            os.startfile(self.output_dir)
        else:
            messagebox.showwarning("Warning", "Output directory does not exist!")

    def close(self):
        self.root.quit()
