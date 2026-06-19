import pandas as pd
import os
import shutil
import sys
from datetime import date
from openpyxl import load_workbook


# -----------------------------
# Reconciliation Logic
# -----------------------------
def reconcile_data(purchase_file, gstr2a_file):
    purchase_df = pd.read_excel(purchase_file)
    gstr2a_df = pd.read_excel(gstr2a_file)

    # Merge based on Invoice No
    merged = pd.merge(
        purchase_df,
        gstr2a_df,
        on="Invoice No",
        how="left",
        suffixes=(" (Books)", " (2A)")
    )

    # Fill blank values
    merged["Taxable Amt (2A)"] = merged["Taxable Amt (2A)"].fillna(0)
    merged["GST (2A)"] = merged["GST (2A)"].fillna(0)

    # Calculate differences
    merged["Tax Diff (₹)"] = (
        merged["Taxable Amt (Books)"] - merged["Taxable Amt (2A)"]
    )

    merged["GST Diff (₹)"] = (
        merged["GST (Books)"] - merged["GST (2A)"]
    )

    # Status
    merged["Status"] = merged.apply(
        lambda x: "Matched"
        if abs(x["Tax Diff (₹)"]) <= 1 and abs(x["GST Diff (₹)"]) <= 1
        else "Mismatch",
        axis=1
    )

    mismatches = merged[merged["Status"] == "Mismatch"].copy()

    # ITC at Risk
    itc_risk = merged[merged["Taxable Amt (2A)"] == 0].copy()
    itc_risk["Taxable Amt (₹)"] = itc_risk["Taxable Amt (Books)"]
    itc_risk["ITC at Risk (₹)"] = itc_risk["GST (Books)"]
    itc_risk["Recommended Action"] = "Only in Books"

    # Summary
    summary = {
        "Total invoices (Books)": (
            len(merged),
            merged["Taxable Amt (Books)"].sum()
        ),
        "Matched": (
            len(merged[merged["Status"] == "Matched"]),
            merged[merged["Status"] == "Matched"]["Taxable Amt (Books)"].sum()
        ),
        "Mismatches": (
            len(mismatches),
            mismatches["Taxable Amt (Books)"].sum()
        ),
        "Missing in 2A": (
            len(itc_risk),
            itc_risk["Taxable Amt (Books)"].sum()
        ),
        "Extra in 2A": (0, 0),
        "ITC safe to claim": (
            "-",
            merged[merged["Status"] == "Matched"]["GST (Books)"].sum()
        ),
        "ITC at risk": (
            "-",
            itc_risk["GST (Books)"].sum()
        ),
    }

    return {
        "summary": summary,
        "full_reconciliation": merged,
        "mismatches": mismatches,
        "itc_risk": itc_risk
    }


# -----------------------------
# Save Report
# -----------------------------
def save_report(result, tax_month):
    template_file = "GST_Reconciliation_Report_April (3).xlsx"

    output_folder = "reports"
    os.makedirs(output_folder, exist_ok=True)

    output_file = os.path.join(
        output_folder,
        f"GST_Reconciliation_Report_{tax_month}.xlsx"
    )

    shutil.copy(template_file, output_file)

    wb = load_workbook(output_file)

    # -----------------------------
    # Summary Sheet
    # -----------------------------
    ws = wb["Summary"]

    ws.merge_cells("A2:C2")
    ws["A2"] = f"Generated on {date.today().strftime('%d %b %Y')} | Tolerance: ₹1.0"

    summary_rows = list(result["summary"].items())

    start_row = 5
    for i, (metric, values) in enumerate(summary_rows):
        ws.cell(start_row + i, 1, metric)
        ws.cell(start_row + i, 2, values[0])
        ws.cell(start_row + i, 3, values[1])

    # -----------------------------
    # Full Reconciliation
    # -----------------------------
    ws = wb["Full Reconciliation"]

    start_row = 3
    for _, row in result["full_reconciliation"].iterrows():
        ws.cell(start_row, 1, row["Invoice No"])
        ws.cell(start_row, 2, row["Vendor Name"])
        ws.cell(start_row, 3, row["GSTIN"])
        ws.cell(start_row, 4, row["Invoice Date"])
        ws.cell(start_row, 5, row["Taxable Amt (Books)"])
        ws.cell(start_row, 6, row["Taxable Amt (2A)"])
        ws.cell(start_row, 7, row["Tax Diff (₹)"])
        ws.cell(start_row, 8, row["GST (Books)"])
        ws.cell(start_row, 9, row["GST (2A)"])
        ws.cell(start_row, 10, row["GST Diff (₹)"])
        ws.cell(start_row, 11, row["Status"])
        start_row += 1

    # -----------------------------
    # Mismatches
    # -----------------------------
    ws = wb["Mismatches"]

    start_row = 3
    for _, row in result["mismatches"].iterrows():
        ws.cell(start_row, 1, row["Invoice No"])
        ws.cell(start_row, 2, row["Vendor Name"])
        ws.cell(start_row, 3, row["GSTIN"])
        ws.cell(start_row, 4, row["Invoice Date"])
        ws.cell(start_row, 5, row["Taxable Amt (Books)"])
        ws.cell(start_row, 6, row["Taxable Amt (2A)"])
        ws.cell(start_row, 7, row["Tax Diff (₹)"])
        ws.cell(start_row, 8, row["GST (Books)"])
        ws.cell(start_row, 9, row["GST (2A)"])
        ws.cell(start_row, 10, row["GST Diff (₹)"])
        ws.cell(start_row, 11, row["Status"])
        start_row += 1

    # -----------------------------
    # ITC at Risk
    # -----------------------------
    ws = wb["ITC at Risk"]

    start_row = 4
    for _, row in result["itc_risk"].iterrows():
        ws.cell(start_row, 1, row["Invoice No"])
        ws.cell(start_row, 2, row["Vendor Name"])
        ws.cell(start_row, 3, row["GSTIN"])
        ws.cell(start_row, 4, row["Invoice Date"])
        ws.cell(start_row, 5, row["Taxable Amt (₹)"])
        ws.cell(start_row, 6, row["ITC at Risk (₹)"])
        ws.cell(start_row, 7, row["Recommended Action"])
        start_row += 1

    wb.save(output_file)

    return output_file


# -----------------------------
# Main Function
# -----------------------------
def main():
    if len(sys.argv) != 4:
        print("Usage: python gst_reconciliation.py purchase.xlsx gstr.xlsx month")
        return

    purchase_file = sys.argv[1]
    gstr2a_file = sys.argv[2]
    tax_month = sys.argv[3]

    print("Loading data...")
    print("Reconciling...")

    result = reconcile_data(purchase_file, gstr2a_file)

    print("Saving report...")
    output_path = save_report(result, tax_month)

    print(f"Report saved at: {output_path}")


if __name__ == "__main__":
    main()