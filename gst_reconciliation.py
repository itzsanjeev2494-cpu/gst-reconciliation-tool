import pandas as pd
import sys
import os

# Read command-line arguments
purchase_file = sys.argv[1]
gstr_file = sys.argv[2]
TAX_MONTH = sys.argv[3]

# Load Excel files
purchase_df = pd.read_excel(purchase_file)
gstr_df = pd.read_excel(gstr_file)

# Standardize column names
purchase_df.columns = purchase_df.columns.str.strip()
gstr_df.columns = gstr_df.columns.str.strip()

# Merge on Invoice No
merged_df = pd.merge(
    purchase_df,
    gstr_df,
    on="Invoice No",
    how="outer",
    suffixes=("_Purchase", "_GSTR2B"),
    indicator=True
)

# Match status logic
def get_status(row):
    if row["_merge"] == "both":
        purchase_taxable = row.get("Taxable Amount_Purchase", 0)
        gstr_taxable = row.get("Taxable Amount_GSTR2B", 0)

        purchase_gst = row.get("GST Amount_Purchase", 0)
        gstr_gst = row.get("GST Amount_GSTR2B", 0)

        if purchase_taxable == gstr_taxable and purchase_gst == gstr_gst:
            return "Matched"
        else:
            return "Mismatch"

    elif row["_merge"] == "left_only":
        return "Missing in GSTR-2B"

    elif row["_merge"] == "right_only":
        return "Missing in Purchase Register"

# Apply reconciliation status
merged_df["Status"] = merged_df.apply(get_status, axis=1)

# Summary sheet
summary = pd.DataFrame({
    "Particulars": [
        "Total Purchase Invoices",
        "Total GSTR-2B Invoices",
        "Matched",
        "Mismatch",
        "Missing in GSTR-2B",
        "Missing in Purchase Register"
    ],
    "Count": [
        len(purchase_df),
        len(gstr_df),
        len(merged_df[merged_df["Status"] == "Matched"]),
        len(merged_df[merged_df["Status"] == "Mismatch"]),
        len(merged_df[merged_df["Status"] == "Missing in GSTR-2B"]),
        len(merged_df[merged_df["Status"] == "Missing in Purchase Register"])
    ]
})

# Create month-wise folder
folder_name = f"{TAX_MONTH}_Recon"
os.makedirs(folder_name, exist_ok=True)

# Final output file path
OUTPUT_FILE = os.path.join(
    folder_name,
    f"GST_Reconciliation_Report_{TAX_MONTH}.xlsx"
)

# Write Excel report
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    merged_df.to_excel(writer, sheet_name="Reconciliation", index=False)
    summary.to_excel(writer, sheet_name="Summary", index=False)

print(f"Reconciliation completed. File saved at: {OUTPUT_FILE}")