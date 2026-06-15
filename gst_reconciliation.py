import pandas as pd
import os
import sys

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

    # Fill missing values
    merged["Taxable Amount_Books"] = merged["Taxable Amount_Books"].fillna(0)
    merged["Taxable Amount_GSTR"] = merged["Taxable Amount_GSTR"].fillna(0)
    merged["GST Amount_Books"] = merged["GST Amount_Books"].fillna(0)
    merged["GST Amount_GSTR"] = merged["GST Amount_GSTR"].fillna(0)

    # Calculate differences
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

    return merged


# ==============================
# SAVE REPORT (MULTI-SHEET)
# ==============================
def save_report(result, tax_month):
    import os
    import shutil
    from openpyxl import load_workbook

    template_file = "GST_Reconciliation_Report_April (3)(3).xlsx"

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

    start_row = 5
    for i, row in summary_data.iterrows():
        ws.cell(row=start_row+i, column=1, value=row["Metric"])
        ws.cell(row=start_row+i, column=2, value=row["Count"])
        ws.cell(row=start_row+i, column=3, value=row["Amount (₹)"])

    # =========================
    # FULL RECONCILIATION
    # =========================
    ws = wb["Full Reconciliation"]

    ws["A1"] = "All Invoices"

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

    ws["A1"] = f"ITC at Risk — {tax_month}"

    for r_idx, row in enumerate(result["itc_risk"].values, start=4):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    wb.save(output_file)

    print(f"Saved: {output_file}")


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