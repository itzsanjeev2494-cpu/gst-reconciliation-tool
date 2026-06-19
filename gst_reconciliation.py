import pandas as pd
import os
import sys
import shutil
from openpyxl import load_workbook

# ==============================
# SETTINGS
# ==============================
TOLERANCE = 1   # Allow ₹1 difference


# ==============================
# LOAD FILES
# ==============================
def load_files(purchase_file, gstr_file):

    books = pd.read_excel(purchase_file)
    gstr2a = pd.read_excel(gstr_file)

    # Standardize columns
    books.columns = books.columns.str.strip()
    gstr2a.columns = gstr2a.columns.str.strip()

    return books, gstr2a


# ==============================
# RECONCILIATION LOGIC
# ==============================
def reconcile(books, gstr2a):

    merged = pd.merge(
        books,
        gstr2a,
        on="Invoice No",
        how="outer",
        suffixes=("_Books", "_GSTR")
    )

    # Fill blanks
    merged["Taxable Amount_Books"] = merged["Taxable Amount_Books"].fillna(0)
    merged["Taxable Amount_GSTR"] = merged["Taxable Amount_GSTR"].fillna(0)
    merged["GST Amount_Books"] = merged["GST Amount_Books"].fillna(0)
    merged["GST Amount_GSTR"] = merged["GST Amount_GSTR"].fillna(0)

    # Difference calculation
    merged["Taxable Difference"] = (
        merged["Taxable Amount_Books"] - merged["Taxable Amount_GSTR"]
    )

    merged["GST Difference"] = (
        merged["GST Amount_Books"] - merged["GST Amount_GSTR"]
    )

    # Status classification
    def classify(row):

        tax_diff = abs(row["Taxable Difference"])
        gst_diff = abs(row["GST Difference"])

        if row["Taxable Amount_Books"] == 0:
            return "Only in GSTR"

        elif row["Taxable Amount_GSTR"] == 0:
            return "Only in Books"

        elif tax_diff <= TOLERANCE and gst_diff <= TOLERANCE:
            return "Matched"

        else:
            return "Mismatch"

    merged["Status"] = merged.apply(classify, axis=1)

    mismatch = merged[merged["Status"] == "Mismatch"]
    itc_risk = merged[merged["Status"] == "Only in Books"]

    # Summary dictionary
    summary = {
        "total_books": len(books),
        "total_books_amount": books["Taxable Amount"].sum(),

        "matched_count": len(merged[merged["Status"] == "Matched"]),
        "matched_amount": merged.loc[
            merged["Status"] == "Matched",
            "Taxable Amount_Books"
        ].sum(),

        "mismatch_count": len(merged[merged["Status"] == "Mismatch"]),
        "mismatch_amount": merged.loc[
            merged["Status"] == "Mismatch",
            "Taxable Amount_Books"
        ].sum(),

        "missing_count": len(merged[merged["Status"] == "Only in Books"]),
        "missing_amount": merged.loc[
            merged["Status"] == "Only in Books",
            "Taxable Amount_Books"
        ].sum(),

        "extra_count": len(merged[merged["Status"] == "Only in GSTR"]),
        "extra_amount": merged.loc[
            merged["Status"] == "Only in GSTR",
            "Taxable Amount_GSTR"
        ].sum(),

        "itc_safe": merged.loc[
            merged["Status"] == "Matched",
            "GST Amount_Books"
        ].sum(),

        "itc_risk": merged.loc[
            merged["Status"] == "Only in Books",
            "GST Amount_Books"
        ].sum()
    }

    return {
        "summary": summary,
        "full": merged,
        "mismatch": mismatch,
        "itc_risk": itc_risk
    }


# ==============================
# SAVE REPORT (TEMPLATE BASED)
# ==============================
def save_report(result, tax_month):

    template_file = "GST_Reconciliation_Report_April (3).xlsx"

    output_folder = f"{tax_month}_Recon"
    os.makedirs(output_folder, exist_ok=True)

    output_file = os.path.join(
        output_folder,
        f"GST_Reconciliation_Report_{tax_month}.xlsx"
    )

    # Copy template
    shutil.copy(template_file, output_file)

    wb = load_workbook(output_file)

    # =========================
    # SUMMARY SHEET
    # =========================
    ws = wb["Summary"]

    ws["A1"] = f"GST Reconciliation Report — {tax_month}"

    summary_data = result["summary"]

    summary_rows = [
        ["Total invoices (Books)", summary_data["total_books"], summary_data["total_books_amount"]],
        ["✓ Matched", summary_data["matched_count"], summary_data["matched_amount"]],
        ["✗ Mismatches", summary_data["mismatch_count"], summary_data["mismatch_amount"]],
        ["⚠ Missing in 2A", summary_data["missing_count"], summary_data["missing_amount"]],
        ["ℹ Extra in 2A", summary_data["extra_count"], summary_data["extra_amount"]],
        ["ITC safe to claim", "-", summary_data["itc_safe"]],
        ["ITC at risk", "-", summary_data["itc_risk"]]
    ]

    for i, row in enumerate(summary_rows, start=5):
        ws[f"A{i}"] = row[0]
        ws[f"B{i}"] = row[1]
        ws[f"C{i}"] = row[2]

    # =========================
    # FULL RECONCILIATION
    # =========================
    ws = wb["Full Reconciliation"]

    for r_idx, row in enumerate(result["full"].values, start=3):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # =========================
    # MISMATCHES
    # =========================
    ws = wb["Mismatches"]

    for r_idx, row in enumerate(result["mismatch"].values, start=3):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # =========================
    # ITC AT RISK
    # =========================
    ws = wb["ITC at Risk"]

    for r_idx, row in enumerate(result["itc_risk"].values, start=4):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    wb.save(output_file)

    print(f"Saved: {output_file}")

    return output_file


# ==============================
# MAIN FUNCTION
# ==============================
def main():

    if len(sys.argv) != 4:
        print("Usage: python gst_reconciliation.py purchase.xlsx gstr.xlsx month")
        sys.exit(1)

    purchase_file = sys.argv[1]
    gstr_file = sys.argv[2]
    tax_month = sys.argv[3]

    print("Loading data...")
    books, gstr2a = load_files(purchase_file, gstr_file)

    print("Reconciling...")
    result = reconcile(books, gstr2a)

    print("Saving report...")
    output_path = save_report(result, tax_month)

    print(f"Report saved at: {output_path}")


if __name__ == "__main__":
    main()