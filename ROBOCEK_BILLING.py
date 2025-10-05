"""
ROBOCEK GCEK Billing GUI
- Python 3.8+
- Tkinter GUI for entering recipient and components (name, qty, unit price)
- Calculates subtotal, discount, tax, grand total using Decimal
- Generates a PDF invoice styled as ROBOCEK acknowledgement form (uses reportlab)
- Saves CSV invoice record (history)
"""

import csv
import os
import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

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
        self.title("ROBOCEK GCEK Billing")
        self.geometry("900x650")
        self.minsize(800, 600)

        # Default vice president name
        self.vp_name = tk.StringVar(value="Harikesh O P")

        self.recipient_name = tk.StringVar()
        self.item_name = tk.StringVar()
        self.item_qty = tk.StringVar(value="1")
        self.item_price = tk.StringVar(value="0.00")
        self.discount_pct = tk.StringVar(value="0")
        self.tax_pct = tk.StringVar(value="18")

        self.items = []  # list of dicts: {"name": ..., "qty": Decimal, "price": Decimal}

        self._build_ui()

    def _build_ui(self):
        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", padx=12, pady=8)

        # Recipient / invoice info
        ttk.Label(frm_top, text="Recipient / Billing Person:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm_top, textvariable=self.recipient_name, width=30).grid(row=0, column=1, sticky="w", padx=(6,20))

        ttk.Label(frm_top, text="Invoice No:").grid(row=0, column=2, sticky="e")
        self.invoice_no_var = tk.StringVar(value=self._gen_invoice_no())
        ttk.Entry(frm_top, textvariable=self.invoice_no_var, width=18).grid(row=0, column=3, sticky="w", padx=(6,20))

        ttk.Label(frm_top, text="Date:").grid(row=0, column=4, sticky="e")
        self.date_var = tk.StringVar(value=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        ttk.Entry(frm_top, textvariable=self.date_var, width=22).grid(row=0, column=5, sticky="w")

        # Vice President name (left-side signature)
        ttk.Label(frm_top, text="Vice President (signature):").grid(row=1, column=0, sticky="w", pady=(8,0))
        ttk.Entry(frm_top, textvariable=self.vp_name, width=30).grid(row=1, column=1, sticky="w", padx=(6,20), pady=(8,0))
        ttk.Label(frm_top, text="Organization:").grid(row=1, column=2, sticky="e", pady=(8,0))
        ttk.Label(frm_top, text="ROBOCEK GCEK", foreground="green").grid(row=1, column=3, sticky="w", padx=(6,20), pady=(8,0))

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
        ttk.Button(btn_frame, text="Generate PDF Invoice", command=self.generate_pdf).pack(side="right", padx=6)
        ttk.Button(btn_frame, text="Reset / New Invoice", command=self.reset_invoice).pack(side="right", padx=6)

        # Bind events
        self.discount_pct.trace_add("write", lambda *a: self._recalc_totals())
        self.tax_pct.trace_add("write", lambda *a: self._recalc_totals())

    def _gen_invoice_no(self):
        return datetime.datetime.now().strftime("%Y%m%d%H%M%S")

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

        story.append(Paragraph("<b>ROBOCEK GCEK Acknowledgement Form</b>", styleCenter))
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"Date: {self.date_var.get()}", styleN))
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"Recipient: {self.recipient_name.get()}", styleN))
        story.append(Spacer(1, 8))

        # Body
        body_text = f"Dear {self.recipient_name.get() or '<n>'},"
        story.append(Paragraph(body_text, styleN))
        story.append(Spacer(1, 6))
        body_text2 = "This letter serves as an acknowledgment of the components rented from ROBOCEK GCEK. The following components have been rented:"
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
        # Left: ROBOCEK GCEK with VP name and title
        # Right: Recipient name and signature line
        footer_data = [
            [
                Paragraph(f"<b>ROBOCEK GCEK</b><br/>{self.vp_name.get()}<br/>Vice President", styleN),
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

    def reset_invoice(self):
        if not messagebox.askyesno("Confirm", "Reset invoice? All current items will be cleared."):
            return
        self.items.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.recipient_name.set("")
        self.discount_pct.set("0")
        self.tax_pct.set("18")
        self.invoice_no_var.set(self._gen_invoice_no())
        self.date_var.set(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self._recalc_totals()


if __name__ == "__main__":
    app = BillingApp()
    app.mainloop()