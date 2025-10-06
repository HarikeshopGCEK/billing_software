"""
IEEE Billing GUI
- Python 3.8+
- Tkinter GUI for entering recipient and components (name, qty, unit price)
- Calculates subtotal, discount, tax, grand total using Decimal
- Generates a PDF invoice styled as IEEE acknowledgement form (uses reportlab)
- Saves CSV invoice record (history)
"""

import csv
import os
import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Excel support (optional)
try:
    import pandas as pd
    excel_support = True
except ImportError:
    excel_support = False

# PDF generation (reportlab)
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    from reportlab.lib import colors
    from reportlab.lib.units import mm
except Exception as e:
    reportlab = None
    # We'll check later and show helpful error if missing
else:
    reportlab = True

# ---------- Utility helpers ----------
def D(v):
    """Convert value to Decimal safely (2 decimal places)."""
    try:
        d = Decimal(str(v))
    except (InvalidOperation, ValueError):
        d = Decimal("0")
    return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

def fmt_money(d):
    return f"{D(d):,.2f}"

# ---------- Main Application ----------
class BillingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("IEEE Billing Software")
        self.geometry("900x650")
        self.minsize(800, 600)

        # Default chairperson name and position
        self.chair_name = tk.StringVar(value="Harikesh O P")
        self.chair_position = tk.StringVar(value="IEEE RAS SB Chairperson")

        self.recipient_name = tk.StringVar()
        self.recipient_phone = tk.StringVar()
        self.item_name = tk.StringVar()
        self.item_qty = tk.StringVar(value="1")
        self.item_price = tk.StringVar(value="0.00")
        self.discount_pct = tk.StringVar(value="0")
        self.tax_pct = tk.StringVar(value="18")

        self.items = []  # list of dicts: {"name": ..., "qty": Decimal, "price": Decimal}
        
        # Master invoice database file
        self.master_db_file = "invoice_database.csv"

        self._build_ui()

    def _build_ui(self):
        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", padx=12, pady=8)

        # Recipient / invoice info
        ttk.Label(frm_top, text="Recipient / Billing Person: *", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Entry(frm_top, textvariable=self.recipient_name, width=25).grid(row=0, column=1, sticky="w", padx=(6,12))

        ttk.Label(frm_top, text="Phone Number: *", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=2, sticky="w", padx=(6,0))
        ttk.Entry(frm_top, textvariable=self.recipient_phone, width=15).grid(row=0, column=3, sticky="w", padx=(6,12))

        ttk.Label(frm_top, text="Invoice No:").grid(row=0, column=4, sticky="e")
        self.invoice_no_var = tk.StringVar(value=self._gen_invoice_no())
        ttk.Entry(frm_top, textvariable=self.invoice_no_var, width=18).grid(row=0, column=5, sticky="w", padx=(6,12))

        ttk.Label(frm_top, text="Date:").grid(row=0, column=6, sticky="e")
        self.date_var = tk.StringVar(value=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        ttk.Entry(frm_top, textvariable=self.date_var, width=22).grid(row=0, column=7, sticky="w")

        # Chairperson name and position
        ttk.Label(frm_top, text="Chairperson Name:").grid(row=1, column=0, sticky="w", pady=(8,0))
        ttk.Entry(frm_top, textvariable=self.chair_name, width=25).grid(row=1, column=1, sticky="w", padx=(6,12), pady=(8,0))
        
        ttk.Label(frm_top, text="Position:").grid(row=1, column=2, sticky="w", padx=(6,0), pady=(8,0))
        ttk.Entry(frm_top, textvariable=self.chair_position, width=20).grid(row=1, column=3, sticky="w", padx=(6,12), pady=(8,0))
        
        ttk.Label(frm_top, text="Organization:").grid(row=1, column=4, sticky="e", pady=(8,0))
        ttk.Label(frm_top, text="IEEE SB GCEK", foreground="blue").grid(row=1, column=5, sticky="w", padx=(6,0), pady=(8,0))

        # Item entry frame
        box = ttk.LabelFrame(self, text="Add Component / Item")
        box.pack(fill="x", padx=12, pady=(10,6))

        ttk.Label(box, text="Item name:").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        ttk.Entry(box, textvariable=self.item_name, width=36).grid(row=0, column=1, padx=6, pady=6, sticky="w")

        ttk.Label(box, text="Qty:").grid(row=0, column=2, padx=6, pady=6, sticky="w")
        ttk.Entry(box, textvariable=self.item_qty, width=8).grid(row=0, column=3, padx=6, pady=6, sticky="w")

        ttk.Label(box, text="Unit price (Rs):").grid(row=0, column=4, padx=6, pady=6, sticky="w")
        ttk.Entry(box, textvariable=self.item_price, width=12).grid(row=0, column=5, padx=6, pady=6, sticky="w")

        ttk.Button(box, text="Add Item", command=self.add_item).grid(row=0, column=6, padx=8, pady=6)
        ttk.Button(box, text="Clear Fields", command=self.clear_item_fields).grid(row=0, column=7, padx=6, pady=6)

        # Items list (Treeview)
        tv_frame = ttk.Frame(self)
        tv_frame.pack(fill="both", expand=True, padx=12, pady=6)

        cols = ("#1", "#2", "#3", "#4")
        self.tree = ttk.Treeview(tv_frame, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("#1", text="Item")
        self.tree.heading("#2", text="Qty")
        self.tree.heading("#3", text="Unit Price (Rs)")
        self.tree.heading("#4", text="Total (Rs)")
        self.tree.column("#1", anchor="w", width=360)
        self.tree.column("#2", anchor="center", width=60)
        self.tree.column("#3", anchor="e", width=120)
        self.tree.column("#4", anchor="e", width=120)
        self.tree.pack(fill="both", expand=True, side="left")

        scrollbar = ttk.Scrollbar(tv_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # Bottom area for totals and actions
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=12, pady=8)

        # Discounts / tax
        ttk.Label(bottom, text="Discount %:").grid(row=0, column=0, sticky="w", padx=6)
        ttk.Entry(bottom, textvariable=self.discount_pct, width=8).grid(row=0, column=1, sticky="w", padx=6)

        ttk.Label(bottom, text="Tax %:").grid(row=0, column=2, sticky="w", padx=6)
        ttk.Entry(bottom, textvariable=self.tax_pct, width=8).grid(row=0, column=3, sticky="w", padx=6)

        # Totals display
        ttk.Label(bottom, text="Subtotal (Rs):").grid(row=1, column=0, sticky="e", padx=6, pady=(8,0))
        self.subtotal_var = tk.StringVar(value="0.00")
        ttk.Label(bottom, textvariable=self.subtotal_var).grid(row=1, column=1, sticky="w", padx=6, pady=(8,0))

        ttk.Label(bottom, text="Discount Amt (Rs):").grid(row=1, column=2, sticky="e", padx=6, pady=(8,0))
        self.discount_amt_var = tk.StringVar(value="0.00")
        ttk.Label(bottom, textvariable=self.discount_amt_var).grid(row=1, column=3, sticky="w", padx=6, pady=(8,0))

        ttk.Label(bottom, text="Tax Amt (Rs):").grid(row=1, column=4, sticky="e", padx=6, pady=(8,0))
        self.tax_amt_var = tk.StringVar(value="0.00")
        ttk.Label(bottom, textvariable=self.tax_amt_var).grid(row=1, column=5, sticky="w", padx=6, pady=(8,0))

        ttk.Label(bottom, text="Grand Total (Rs):", font=("TkDefaultFont", 10, "bold")).grid(row=2, column=0, sticky="e", padx=6, pady=(12,0))
        self.grand_total_var = tk.StringVar(value="0.00")
        ttk.Label(bottom, textvariable=self.grand_total_var, font=("TkDefaultFont", 10, "bold")).grid(row=2, column=1, sticky="w", padx=6, pady=(12,0))

        # Action buttons
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=12, pady=(6,12))

        ttk.Button(btn_frame, text="Remove Selected Item", command=self.remove_selected_item).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Save CSV (history)", command=self.save_csv).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="View Invoice History", command=self.view_invoice_history).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Generate PDF Invoice", command=self.generate_pdf).pack(side="right", padx=6)
        ttk.Button(btn_frame, text="Reset / New Invoice", command=self.reset_invoice).pack(side="right", padx=6)

        # Bind events
        self.discount_pct.trace_add("write", lambda *a: self._recalc_totals())
        self.tax_pct.trace_add("write", lambda *a: self._recalc_totals())

    def _gen_invoice_no(self):
        """Generate next invoice number by incrementing from the last used number"""
        # Try to get the last invoice number from the database
        if os.path.exists(self.master_db_file):
            try:
                with open(self.master_db_file, "r", encoding="utf-8") as csvfile:
                    reader = csv.DictReader(csvfile)
                    last_invoice = None
                    for row in reader:
                        last_invoice = row.get("Invoice No", "")
                    
                    if last_invoice:
                        # Extract the number part and increment
                        try:
                            # Assuming invoice numbers are in format like "20241201123456" (timestamp)
                            # or simple numbers like "1", "2", "3"
                            if last_invoice.isdigit():
                                # Simple number format
                                next_num = int(last_invoice) + 1
                                return str(next_num)
                            else:
                                # Timestamp format - just use current timestamp
                                return datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                        except ValueError:
                            # If parsing fails, use current timestamp
                            return datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            except Exception:
                pass
        
        # If no database exists or error occurred, start with 1
        return "1"

    def clear_item_fields(self):
        self.item_name.set("")
        self.item_qty.set("1")
        self.item_price.set("0.00")

    def add_item(self):
        name = self.item_name.get().strip()
        if not name:
            messagebox.showwarning("Validation", "Item name cannot be empty.")
            return
        try:
            qty = D(self.item_qty.get())
            price = D(self.item_price.get())
        except Exception:
            messagebox.showwarning("Validation", "Quantity and price must be numeric.")
            return
        if qty <= 0:
            messagebox.showwarning("Validation", "Quantity must be greater than 0.")
            return
        if price < 0:
            messagebox.showwarning("Validation", "Price cannot be negative.")
            return
        total = (qty * price).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        self.items.append({"name": name, "qty": qty, "price": price, "total": total})
        self.tree.insert("", "end", values=(name, f"{qty.normalize()}", f"{price:.2f}", f"{total:.2f}"))
        self.clear_item_fields()
        self._recalc_totals()

    def remove_selected_item(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Info", "No item selected.")
            return
        idx = self.tree.index(sel[0])
        self.tree.delete(sel[0])
        try:
            del self.items[idx]
        except Exception:
            pass
        self._recalc_totals()

    def _recalc_totals(self):
        subtotal = sum(it["total"] for it in self.items) if self.items else Decimal("0.00")
        self.subtotal_var.set(fmt_money(subtotal))
        # discount
        try:
            discount_pct = D(self.discount_pct.get())
        except Exception:
            discount_pct = Decimal("0")
        discount_amt = (subtotal * discount_pct / Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        self.discount_amt_var.set(fmt_money(discount_amt))
        # tax
        try:
            tax_pct = D(self.tax_pct.get())
        except Exception:
            tax_pct = Decimal("0")
        taxable = (subtotal - discount_amt).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        tax_amt = (taxable * tax_pct / Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        self.tax_amt_var.set(fmt_money(tax_amt))
        grand = (taxable + tax_amt).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        self.grand_total_var.set(fmt_money(grand))

    def save_csv(self):
        if not self.items:
            messagebox.showwarning("No items", "Add at least one item before saving CSV.")
            return
        
        # Validate required fields
        if not self.recipient_name.get().strip():
            messagebox.showerror("Required Field Missing", "Recipient name is required. Please enter the recipient's name.")
            return
        
        if not self.recipient_phone.get().strip():
            messagebox.showerror("Required Field Missing", "Phone number is required. Please enter the recipient's phone number.")
            return
        folder = filedialog.askdirectory(title="Choose folder to save CSV")
        if not folder:
            return
        fname = f"invoice_{self.invoice_no_var.get()}.csv"
        path = os.path.join(folder, fname)
        with open(path, "w", newline="") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["Invoice No", self.invoice_no_var.get()])
            writer.writerow(["Date", self.date_var.get()])
            writer.writerow(["Recipient", self.recipient_name.get()])
            writer.writerow(["Phone Number", self.recipient_phone.get()])
            writer.writerow([])
            writer.writerow(["Item", "Qty", "Unit Price", "Total"])
            for it in self.items:
                writer.writerow([it["name"], str(it["qty"]), f"{it['price']:.2f}", f"{it['total']:.2f}"])
            writer.writerow([])
            writer.writerow(["Subtotal", self.subtotal_var.get()])
            writer.writerow([f"Discount ({self.discount_pct.get()}%)", self.discount_amt_var.get()])
            writer.writerow([f"Tax ({self.tax_pct.get()}%)", self.tax_amt_var.get()])
            writer.writerow(["Grand Total", self.grand_total_var.get()])

        messagebox.showinfo("Saved", f"CSV saved to:\n{path}")

    def generate_pdf(self):
        if not reportlab:
            messagebox.showerror("Missing dependency",
                                 "reportlab is required to generate PDFs. Install it with:\n\npip install reportlab")
            return
        if not self.items:
            messagebox.showwarning("No items", "Add at least one item before generating PDF.")
            return
        
        # Validate required fields
        if not self.recipient_name.get().strip():
            messagebox.showerror("Required Field Missing", "Recipient name is required. Please enter the recipient's name.")
            return
        
        if not self.recipient_phone.get().strip():
            messagebox.showerror("Required Field Missing", "Phone number is required. Please enter the recipient's phone number.")
            return
        # Ask where to save
        fpath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                             initialfile=f"invoice_{self.invoice_no_var.get()}.pdf")
        if not fpath:
            return

        # Build PDF
        doc = SimpleDocTemplate(fpath, pagesize=A4, topMargin=30, bottomMargin=30, leftMargin=36, rightMargin=36)
        styles = getSampleStyleSheet()
        styleN = styles["Normal"]
        styleH = styles["Heading1"]
        styleCenter = ParagraphStyle(name="Center", parent=styles["Normal"], alignment=TA_CENTER, fontSize=14)
        styleRight = ParagraphStyle(name="Right", parent=styles["Normal"], alignment=TA_RIGHT)
        story = []

        story.append(Paragraph("<b>IEEE Acknowledgement Form</b>", styleCenter))
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"Date: {self.date_var.get()}", styleN))
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"Recipient: {self.recipient_name.get()}", styleN))
        story.append(Paragraph(f"Phone Number: {self.recipient_phone.get()}", styleN))
        story.append(Spacer(1, 8))

        # Body
        body_text = f"Dear {self.recipient_name.get() or '<Name>'},"
        story.append(Paragraph(body_text, styleN))
        story.append(Spacer(1, 6))
        body_text2 = "This letter serves as an acknowledgment of the components rented from IEEE SB GCEK. The following components have been rented:"
        story.append(Paragraph(body_text2, styleN))
        story.append(Spacer(1, 6))

        # Items table
        data = [["Item", "Qty", "Unit Price (Rs)", "Total (Rs)"]]
        for it in self.items:
            data.append([it["name"], str(it["qty"]), f"{it['price']:.2f}", f"{it['total']:.2f}"])
        # summary rows
        data.append(["", "", "Subtotal (Rs):", self.subtotal_var.get()])
        data.append(["", "", f"Discount ({self.discount_pct.get()}%):", self.discount_amt_var.get()])
        data.append(["", "", f"Tax ({self.tax_pct.get()}%):", self.tax_amt_var.get()])
        data.append(["", "", "Grand Total (Rs):", self.grand_total_var.get()])

        t = Table(data, colWidths=[260, 60, 110, 110], hAlign="LEFT")
        t.setStyle(TableStyle([
            ("GRID", (0,0), (-1, -5), 0.5, colors.grey),
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("ALIGN", (1,1), (-1,-1), "CENTER"),
            ("ALIGN", (2,1), (2,-1), "RIGHT"),
            ("ALIGN", (3,1), (3,-1), "RIGHT"),
            ("SPAN", (0, len(data)-4), (2, len(data)-4)),
        ]))
        story.append(t)
        story.append(Spacer(1, 12))

        # Footer with left and right signature
        # Left: IEEE SB GCEK with chair name and title
        # Right: Recipient name and signature line
        footer_data = [
            [
                Paragraph(f"<b>IEEE SB GCEK</b><br/>{self.chair_name.get()}<br/>{self.chair_position.get()}", styleN),
                Paragraph(f"<b>{self.recipient_name.get() or 'Recipient'}</b><br/>Name & Signature", styleRight)
            ]
        ]
        ft = Table(footer_data, colWidths=[270, 270])
        ft.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP")]))
        story.append(ft)

        # Note
        story.append(Spacer(1, 12))
        story.append(Paragraph("Note: Please return rented components in original condition. Any damages or loss incurred "
                               "during the rental period will be your responsibility.", styleN))
        doc.build(story)

        messagebox.showinfo("PDF Generated", f"PDF invoice has been saved to:\n{fpath}")
        
        # Automatically save to master database
        self.save_to_master_db()

    def save_to_master_db(self):
        """Save current invoice to master database CSV file"""
        if not self.items:
            return
            
        # Create master database if it doesn't exist
        file_exists = os.path.exists(self.master_db_file)
        
        with open(self.master_db_file, "a", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            
            # Write header if file is new
            if not file_exists:
                writer.writerow([
                "Invoice No", "Date", "Recipient", "Phone Number", 
                "Items Count", "Subtotal", "Discount %", "Discount Amount", 
                "Tax %", "Tax Amount", "Grand Total", "Chairperson", "Chair Position"
            ])
            
            # Prepare items summary
            items_summary = "; ".join([f"{item['name']} (Qty: {item['qty']}, Price: {item['price']})" 
                                     for item in self.items])
            
            # Write invoice data
            writer.writerow([
                self.invoice_no_var.get(),
                self.date_var.get(),
                self.recipient_name.get(),
                self.recipient_phone.get(),
                len(self.items),
                self.subtotal_var.get(),
                self.discount_pct.get(),
                self.discount_amt_var.get(),
                self.tax_pct.get(),
                self.tax_amt_var.get(),
                self.grand_total_var.get(),
                self.chair_name.get(),
                self.chair_position.get()
            ])

    def view_invoice_history(self):
        """Display invoice history in a new window"""
        if not os.path.exists(self.master_db_file):
            messagebox.showinfo("No History", "No invoice history found. Generate some invoices first!")
            return
            
        # Create new window for history
        history_window = tk.Toplevel(self)
        history_window.title("Invoice History - IEEE Billing")
        history_window.geometry("1000x600")
        
        # Create treeview for history
        frame = ttk.Frame(history_window)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Define columns
        columns = ("Invoice No", "Date", "Recipient", "Phone", "Items", "Subtotal", "Discount", "Tax", "Grand Total")
        tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        
        # Configure columns
        tree.heading("Invoice No", text="Invoice No")
        tree.heading("Date", text="Date")
        tree.heading("Recipient", text="Recipient")
        tree.heading("Phone", text="Phone")
        tree.heading("Items", text="Items Count")
        tree.heading("Subtotal", text="Subtotal (Rs)")
        tree.heading("Discount", text="Discount (Rs)")
        tree.heading("Tax", text="Tax (Rs)")
        tree.heading("Grand Total", text="Grand Total (Rs)")
        
        # Set column widths
        tree.column("Invoice No", width=120, anchor="center")
        tree.column("Date", width=140, anchor="center")
        tree.column("Recipient", width=150, anchor="w")
        tree.column("Phone", width=100, anchor="center")
        tree.column("Items", width=80, anchor="center")
        tree.column("Subtotal", width=100, anchor="e")
        tree.column("Discount", width=100, anchor="e")
        tree.column("Tax", width=100, anchor="e")
        tree.column("Grand Total", width=120, anchor="e")
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        
        # Pack widgets
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Load data
        try:
            with open(self.master_db_file, "r", encoding="utf-8") as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    tree.insert("", "end", values=(
                        row.get("Invoice No", ""),
                        row.get("Date", ""),
                        row.get("Recipient", ""),
                        row.get("Phone Number", ""),
                        row.get("Items Count", ""),
                        row.get("Subtotal", ""),
                        row.get("Discount Amount", ""),
                        row.get("Tax Amount", ""),
                        row.get("Grand Total", "")
                    ))
        except Exception as e:
            messagebox.showerror("Error", f"Error reading invoice history: {str(e)}")
            return
        
        # Add buttons frame
        btn_frame = ttk.Frame(history_window)
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(btn_frame, text="Export to Excel", command=lambda: self.export_history_to_excel()).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Refresh", command=lambda: self.refresh_history(tree)).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Close", command=history_window.destroy).pack(side="right", padx=5)
        
        # Store reference for refresh
        self.history_tree = tree

    def refresh_history(self, tree):
        """Refresh the history treeview"""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
            
        # Reload data
        if os.path.exists(self.master_db_file):
            try:
                with open(self.master_db_file, "r", encoding="utf-8") as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        tree.insert("", "end", values=(
                            row.get("Invoice No", ""),
                            row.get("Date", ""),
                            row.get("Recipient", ""),
                            row.get("Phone Number", ""),
                            row.get("Items Count", ""),
                            row.get("Subtotal", ""),
                            row.get("Discount Amount", ""),
                            row.get("Tax Amount", ""),
                            row.get("Grand Total", "")
                        ))
            except Exception as e:
                messagebox.showerror("Error", f"Error refreshing history: {str(e)}")

    def export_history_to_excel(self):
        """Export invoice history to Excel file"""
        if not excel_support:
            messagebox.showerror("Excel Support Missing", 
                               "pandas library is required for Excel export.\nInstall it with: pip install pandas openpyxl")
            return
            
        if not os.path.exists(self.master_db_file):
            messagebox.showinfo("No History", "No invoice history found to export!")
            return
            
        # Ask where to save Excel file
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"invoice_history_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        
        if not file_path:
            return
            
        try:
            # Read CSV and convert to Excel
            df = pd.read_csv(self.master_db_file)
            
            # Create Excel writer with formatting
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Invoice History', index=False)
                
                # Get the workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets['Invoice History']
                
                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            messagebox.showinfo("Export Successful", f"Invoice history exported to:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Error exporting to Excel: {str(e)}")

    def reset_invoice(self):
        if not messagebox.askyesno("Confirm", "Reset invoice? All current items will be cleared."):
            return
        self.items.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.recipient_name.set("")
        self.recipient_phone.set("")
        self.chair_position.set("IEEE RAS SB Chairperson")  # Reset to default
        self.discount_pct.set("0")
        self.tax_pct.set("18")
        self.invoice_no_var.set(self._gen_invoice_no())  # Auto-increment invoice number
        self.date_var.set(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self._recalc_totals()


if __name__ == "__main__":
    app = BillingApp()
    app.mainloop()