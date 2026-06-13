"""
GST Reconciliation Tool — GSTR-2A vs Purchase Register
========================================================
Author  : [Your Name], CA Finalist
Purpose : Automates monthly GST ITC reconciliation
Usage   : python gst_reconciliation.py
Inputs  : purchase_register.xlsx, gstr_2a.xlsx (in same folder)
Output  : GST_Reconciliation_Report.xlsx
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment,
                              Border, Side, numbers)
from openpyxl.utils import get_column_letter
from datetime import datetime

# ── 1. CONFIG ────────────────────────────────────────────────────────────────

import sys
import os

BOOKS_FILE = sys.argv[1]
GSTR2A_FILE = sys.argv[2]
TAX_MONTH = sys.argv[3]

TOLERANCE = 1.0

folder_name = TAX_MONTH + "_Recon"

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

OUTPUT_FILE = os.path.join(
    folder_name,
    f"GST_Reconciliation_Report_{TAX_MONTH}.xlsx"
)


# ── 2. COLOURS (openpyxl hex, no leading #) ───────────────────────────────

C = {
    "header_bg"  : "1F3864",   # dark navy
    "header_fg"  : "FFFFFF",
    "matched_bg" : "EAF3DE",   # green tint
    "matched_fg" : "276221",
    "mismatch_bg": "FCEBEB",   # red tint
    "mismatch_fg": "7B1A1A",
    "missing_bg" : "FAEEDA",   # amber tint
    "missing_fg" : "6B3D0B",
    "extra_bg"   : "E6F1FB",   # blue tint
    "extra_fg"   : "0C447C",
    "summary_bg" : "F2F2F2",
    "title_bg"   : "D6E4F0",
    "row_alt"    : "F9F9F9",
}

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size, name="Arial")

def border_thin():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

# ── 3. LOAD DATA ─────────────────────────────────────────────────────────────

def load_data():
    books  = pd.read_excel(BOOKS_FILE,  dtype={"Invoice No": str, "GSTIN": str})
    gstr2a = pd.read_excel(GSTR2A_FILE, dtype={"Invoice No": str, "GSTIN": str})

    # Normalise column names
    books.columns  = books.columns.str.strip()
    gstr2a.columns = gstr2a.columns.str.strip()

    # Normalise keys for matching
    for df in [books, gstr2a]:
        df["Invoice No"] = df["Invoice No"].str.upper().str.strip()
        df["GSTIN"]      = df["GSTIN"].str.upper().str.strip()

    return books, gstr2a

# ── 4. RECONCILIATION LOGIC ───────────────────────────────────────────────

def reconcile(books, gstr2a):
    merged = books.merge(
        gstr2a[["Invoice No", "GSTIN", "Taxable Amount", "GST Amount"]],
        on=["Invoice No", "GSTIN"],
        how="outer",
        suffixes=("_Books", "_2A"),
        indicator=True
    )

    def classify(row):
        if row["_merge"] == "left_only":
            return "Missing in 2A"
        if row["_merge"] == "right_only":
            return "Extra in 2A"
        tax_diff = abs(row["Taxable Amount_Books"] - row["Taxable Amount_2A"])
        gst_diff = abs(row["GST Amount_Books"]     - row["GST Amount_2A"])
        if tax_diff > TOLERANCE or gst_diff > TOLERANCE:
            return "Mismatch"
        return "Matched"

    merged["Status"]        = merged.apply(classify, axis=1)
    merged["Tax Diff (₹)"]  = merged["Taxable Amount_2A"].fillna(0) - merged["Taxable Amount_Books"].fillna(0)
    merged["GST Diff (₹)"]  = merged["GST Amount_2A"].fillna(0)     - merged["GST Amount_Books"].fillna(0)
    merged["ITC at Risk (₹)"] = merged.apply(
        lambda r: r["GST Amount_Books"] if r["Status"] == "Missing in 2A" else 0, axis=1
    )

    return merged

# ── 5. BUILD EXCEL REPORT ────────────────────────────────────────────────────

def style_header_row(ws, row_num, cols):
    for col in range(1, cols + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill      = fill(C["header_bg"])
        cell.font      = font(bold=True, color=C["header_fg"], size=10)
        cell.alignment = center()
        cell.border    = border_thin()

def status_colors(status):
    mapping = {
        "Matched"      : (C["matched_bg"],  C["matched_fg"]),
        "Mismatch"     : (C["mismatch_bg"], C["mismatch_fg"]),
        "Missing in 2A": (C["missing_bg"],  C["missing_fg"]),
        "Extra in 2A"  : (C["extra_bg"],    C["extra_fg"]),
    }
    return mapping.get(status, ("FFFFFF", "000000"))

def write_summary(ws, df):
    total     = len(df[df["_merge"] != "right_only"])
    matched   = len(df[df["Status"] == "Matched"])
    mismatch  = len(df[df["Status"] == "Mismatch"])
    missing   = len(df[df["Status"] == "Missing in 2A"])
    extra     = len(df[df["Status"] == "Extra in 2A"])
    itc_risk  = df["ITC at Risk (₹)"].sum()
    itc_safe  = df.loc[df["Status"] == "Matched", "GST Amount_Books"].sum()

    ws.merge_cells("A1:H1")
    title = ws["A1"]
    title.value     = f"GST Reconciliation Report — {TAX_MONTH}"
    title.font      = font(bold=True, color="1F3864", size=14)
    title.alignment = center()
    title.fill      = fill(C["title_bg"])

    ws.merge_cells("A2:H2")
    ws["A2"].value     = f"Generated on {datetime.today().strftime('%d %b %Y')}  |  Tolerance: ₹{TOLERANCE}"
    ws["A2"].font      = font(color="555555", size=9)
    ws["A2"].alignment = center()

    headers = ["Metric", "Count", "Amount (₹)"]
    rows = [
        ("Total invoices (Books)", total,    df["Taxable Amount_Books"].sum()),
        ("✔  Matched",             matched,  df.loc[df["Status"]=="Matched","Taxable Amount_Books"].sum()),
        ("✖  Mismatches",          mismatch, df.loc[df["Status"]=="Mismatch","Taxable Amount_Books"].sum()),
        ("⚠  Missing in 2A",       missing,  df.loc[df["Status"]=="Missing in 2A","Taxable Amount_Books"].sum()),
        ("ℹ  Extra in 2A",         extra,    df.loc[df["Status"]=="Extra in 2A","Taxable Amount_2A"].sum()),
        ("ITC safe to claim",      "",       itc_safe),
        ("ITC at risk",            "",       itc_risk),
    ]

    for col, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=col)
        c.value = h; c.fill = fill(C["header_bg"])
        c.font = font(bold=True, color="FFFFFF"); c.alignment = center(); c.border = border_thin()

    for i, (label, count, amount) in enumerate(rows, 5):
        bg = C["matched_bg"] if "Matched" in label else \
             C["mismatch_bg"] if "Mismatch" in label else \
             C["missing_bg"] if "Missing" in label else \
             C["extra_bg"] if "Extra" in label else \
             "EBF5E9" if "safe" in label else "FCEBEB" if "risk" in label else C["summary_bg"]
        for col in range(1, 4):
            cell = ws.cell(row=i, column=col); cell.border = border_thin()
            cell.fill = fill(bg)
        ws.cell(row=i, column=1).value = label
        ws.cell(row=i, column=1).font  = font(bold=True, size=10)
        ws.cell(row=i, column=2).value = count if count != "" else "—"
        ws.cell(row=i, column=2).alignment = center()
        ws.cell(row=i, column=3).value = round(amount, 2)
        ws.cell(row=i, column=3).number_format = '₹#,##0.00'
        ws.cell(row=i, column=3).alignment = Alignment(horizontal="right", vertical="center")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 18
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[4].height = 22

def write_detail(ws, df, status_filter=None, sheet_title="Detail"):
    title_text = sheet_title if not status_filter else f"{status_filter} — {sheet_title}"
    data = df if not status_filter else df[df["Status"] == status_filter]

    ws.merge_cells("A1:K1")
    ws["A1"].value = title_text; ws["A1"].font = font(bold=True, color="1F3864", size=12)
    ws["A1"].fill = fill(C["title_bg"]); ws["A1"].alignment = center()

    cols = ["Invoice No", "Vendor Name", "GSTIN", "Invoice Date",
            "Taxable Amt (Books)", "Taxable Amt (2A)", "Tax Diff (₹)",
            "GST (Books)", "GST (2A)", "GST Diff (₹)", "Status"]

    for col, h in enumerate(cols, 1):
        c = ws.cell(row=2, column=col)
        c.value = h; c.fill = fill(C["header_bg"])
        c.font = font(bold=True, color="FFFFFF", size=9)
        c.alignment = center(); c.border = border_thin()

    col_map = {
        "Invoice No"           : "Invoice No",
        "Vendor Name"          : "Vendor Name",
        "GSTIN"                : "GSTIN",
        "Invoice Date"         : "Invoice Date",
        "Taxable Amt (Books)"  : "Taxable Amount_Books",
        "Taxable Amt (2A)"     : "Taxable Amount_2A",
        "Tax Diff (₹)"         : "Tax Diff (₹)",
        "GST (Books)"          : "GST Amount_Books",
        "GST (2A)"             : "GST Amount_2A",
        "GST Diff (₹)"         : "GST Diff (₹)",
        "Status"               : "Status",
    }
    money_cols = {5,6,7,8,9,10}

    for r_idx, (_, row) in enumerate(data.iterrows(), 3):
        bg_row = C["row_alt"] if r_idx % 2 == 0 else "FFFFFF"
        status = row["Status"]
        bg_status, fg_status = status_colors(status)

        for c_idx, col_label in enumerate(cols, 1):
            src = col_map[col_label]
            val = row.get(src, "")
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = border_thin()

            if col_label == "Status":
                cell.value = val; cell.fill = fill(bg_status)
                cell.font = font(bold=True, color=fg_status, size=9)
                cell.alignment = center()
            else:
                cell.value = val if pd.notna(val) else "—"
                cell.fill  = fill(bg_row)
                cell.font  = font(size=9)
                cell.alignment = center() if c_idx in money_cols else left()
                if c_idx in money_cols and pd.notna(val) and val != "—":
                    cell.number_format = '₹#,##0.00'

    widths = [14, 22, 22, 13, 16, 14, 13, 12, 12, 13, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 30
    ws.freeze_panes = "A3"

def write_itc_risk(ws, df):
    missing = df[df["Status"] == "Missing in 2A"].copy()

    ws.merge_cells("A1:F1")
    ws["A1"].value = f"ITC at Risk — {TAX_MONTH}"
    ws["A1"].font  = font(bold=True, color="7B1A1A", size=12)
    ws["A1"].fill  = fill(C["mismatch_bg"]); ws["A1"].alignment = center()

    note_row = 2
    ws.merge_cells(f"A{note_row}:F{note_row}")
    ws[f"A{note_row}"].value = ("⚠ These invoices are not in GSTR-2A. "
                                "ITC cannot be claimed under Rule 36(4) until vendor files GSTR-1.")
    ws[f"A{note_row}"].font  = font(color="6B3D0B", size=9)
    ws[f"A{note_row}"].fill  = fill(C["missing_bg"])
    ws[f"A{note_row}"].alignment = left()

    hdrs = ["Invoice No", "Vendor Name", "GSTIN", "Taxable Amt (₹)", "ITC at Risk (₹)", "Recommended Action"]
    for col, h in enumerate(hdrs, 1):
        c = ws.cell(row=3, column=col)
        c.value = h; c.fill = fill(C["header_bg"])
        c.font = font(bold=True, color="FFFFFF", size=9)
        c.alignment = center(); c.border = border_thin()

    actions = "Follow up with vendor to file GSTR-1; hold ITC claim till reflected in GSTR-2B"
    for r_idx, (_, row) in enumerate(missing.iterrows(), 4):
        bg = C["row_alt"] if r_idx % 2 == 0 else "FFFFFF"
        vals = [row.get("Invoice No",""), row.get("Vendor Name",""),
                row.get("GSTIN",""), row.get("Taxable Amount_Books", 0),
                row.get("ITC at Risk (₹)", 0), actions]
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.value = val if pd.notna(val) else "—"
            cell.fill  = fill(bg); cell.border = border_thin()
            cell.font  = font(size=9)
            if c_idx in {4, 5}:
                cell.number_format = '₹#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = left()

    # Total row
    total_row = 4 + len(missing)
    ws.cell(total_row, 1).value = "TOTAL"; ws.cell(total_row, 1).font = font(bold=True, size=9)
    ws.cell(total_row, 4).value = f"=SUM(D4:D{total_row-1})"; ws.cell(total_row, 4).number_format = '₹#,##0.00'
    ws.cell(total_row, 5).value = f"=SUM(E4:E{total_row-1})"; ws.cell(total_row, 5).number_format = '₹#,##0.00'
    for col in range(1, 7):
        ws.cell(total_row, col).fill   = fill(C["summary_bg"])
        ws.cell(total_row, col).border = border_thin()
        ws.cell(total_row, col).font   = font(bold=True, size=9)
        ws.cell(total_row, col).alignment = Alignment(horizontal="right", vertical="center")

    for i, w in enumerate([14, 22, 22, 16, 16, 52], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    ws.row_dimensions[2].height = 28

# ── 6. MAIN ───────────────────────────────────────────────────────────────────

def main():
    print("Loading data...")
    books, gstr2a = load_data()

    print("Reconciling...")
    result = reconcile(books, gstr2a)

    wb = Workbook()

    ws_summary = wb.active
    ws_summary.title = "Summary"
    write_summary(ws_summary, result)

    ws_all = wb.create_sheet("Full Reconciliation")
    write_detail(ws_all, result, sheet_title="All Invoices")

    ws_mm = wb.create_sheet("Mismatches")
    write_detail(ws_mm, result, status_filter="Mismatch", sheet_title="Mismatches")

    ws_itc = wb.create_sheet("ITC at Risk")
    write_itc_risk(ws_itc, result)

    wb.save(OUTPUT_FILE)
    print(f"\n✅ Report saved: {OUTPUT_FILE}")
    print(f"   Matched       : {len(result[result['Status']=='Matched'])}")
    print(f"   Mismatches    : {len(result[result['Status']=='Mismatch'])}")
    print(f"   Missing in 2A : {len(result[result['Status']=='Missing in 2A'])}")
    print(f"   Extra in 2A   : {len(result[result['Status']=='Extra in 2A'])}")
    print(f"   ITC at risk   : ₹{result['ITC at Risk (₹)'].sum():,.2f}")

if __name__ == "__main__":
    main()
