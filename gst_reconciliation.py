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

    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    folder_name = f"{tax_month}_Recon"

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    output_path = os.path.join(
        folder_name,
        f"GST_Reconciliation_Report_{tax_month}.xlsx"
    )

    summary = result["Status"].value_counts().reset_index()
    summary.columns = ["Status", "Count"]

    mismatches = result[result["Status"] == "Mismatch"]

    itc_risk = result[
        (result["Status"] == "Only in GSTR") |
        (result["Status"] == "Only in Books")
    ]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        summary.to_excel(writer, sheet_name="Summary", index=False)
        result.to_excel(writer, sheet_name="Full Reconciliation", index=False)
        mismatches.to_excel(writer, sheet_name="Mismatches", index=False)
        itc_risk.to_excel(writer, sheet_name="ITC at Risk", index=False)

        wb = writer.book

        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        for ws in wb.worksheets:

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")

            for row in ws.iter_rows():
                for cell in row:
                    cell.border = thin_border

            for col in ws.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)

                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass

                ws.column_dimensions[col_letter].width = max_length + 3

            ws.freeze_panes = "A2"

    return output_path


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