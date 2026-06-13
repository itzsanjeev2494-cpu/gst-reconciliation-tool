import pandas as pd

books_data = [
    {"Invoice No": "INV-001", "Vendor Name": "Tata Consultancy Services Ltd",    "GSTIN": "27AABCT0942C1Z9", "Invoice Date": "01-Apr-2025", "Taxable Amount": 150000, "GST Amount": 27000},
    {"Invoice No": "INV-002", "Vendor Name": "Infosys BPM Limited",              "GSTIN": "29AABCI5570H1ZR", "Invoice Date": "03-Apr-2025", "Taxable Amount":  85000, "GST Amount": 15300},
    {"Invoice No": "INV-003", "Vendor Name": "Reliance Jio Infocomm Ltd",        "GSTIN": "27AAJCR5552K1Z6", "Invoice Date": "05-Apr-2025", "Taxable Amount":  12000, "GST Amount":  2160},
    {"Invoice No": "INV-004", "Vendor Name": "Amazon Web Services India Pvt Ltd","GSTIN": "29AADCA4849H1ZW", "Invoice Date": "07-Apr-2025", "Taxable Amount":  48000, "GST Amount":  8640},
    {"Invoice No": "INV-005", "Vendor Name": "Wipro Limited",                    "GSTIN": "29AAACW0035A1ZM", "Invoice Date": "08-Apr-2025", "Taxable Amount": 200000, "GST Amount": 36000},
    {"Invoice No": "INV-006", "Vendor Name": "HCL Technologies Ltd",             "GSTIN": "09AAACH7718R1ZB", "Invoice Date": "10-Apr-2025", "Taxable Amount":  75000, "GST Amount": 13500},
    {"Invoice No": "INV-007", "Vendor Name": "Mahindra & Mahindra Ltd",          "GSTIN": "27AAACM3025E1ZN", "Invoice Date": "12-Apr-2025", "Taxable Amount":  32000, "GST Amount":  5760},
    {"Invoice No": "INV-008", "Vendor Name": "HDFC Bank Limited",                "GSTIN": "27AAHCH4118A4ZN", "Invoice Date": "14-Apr-2025", "Taxable Amount":  18000, "GST Amount":  3240},
    {"Invoice No": "INV-009", "Vendor Name": "Siemens Limited",                  "GSTIN": "27AAACS4131J1Z2", "Invoice Date": "15-Apr-2025", "Taxable Amount":  60000, "GST Amount": 10800},
    {"Invoice No": "INV-010", "Vendor Name": "Oracle India Pvt Ltd",             "GSTIN": "07AAACO0096F1ZZ", "Invoice Date": "18-Apr-2025", "Taxable Amount":  95000, "GST Amount": 17100},
    {"Invoice No": "INV-011", "Vendor Name": "3M India Limited",                 "GSTIN": "29AAACM0020K1Z8", "Invoice Date": "20-Apr-2025", "Taxable Amount":  42000, "GST Amount":  7560},
    {"Invoice No": "INV-012", "Vendor Name": "Bajaj Electricals Ltd",            "GSTIN": "27AAACB4252E1ZM", "Invoice Date": "22-Apr-2025", "Taxable Amount":  28500, "GST Amount":  5130},
]

# GSTR-2A: same invoices but with deliberate mismatches
gstr2a_data = [
    {"Invoice No": "INV-001", "Vendor Name": "Tata Consultancy Services Ltd",    "GSTIN": "27AABCT0942C1Z9", "Invoice Date": "01-Apr-2025", "Taxable Amount": 150000, "GST Amount": 27000},  # Matched
    {"Invoice No": "INV-002", "Vendor Name": "Infosys BPM Limited",              "GSTIN": "29AABCI5570H1ZR", "Invoice Date": "03-Apr-2025", "Taxable Amount":  85000, "GST Amount": 15300},  # Matched
    {"Invoice No": "INV-003", "Vendor Name": "Reliance Jio Infocomm Ltd",        "GSTIN": "27AAJCR5552K1Z6", "Invoice Date": "05-Apr-2025", "Taxable Amount":  11500, "GST Amount":  2070},  # Taxable mismatch
    {"Invoice No": "INV-004", "Vendor Name": "Amazon Web Services India Pvt Ltd","GSTIN": "29AADCA4849H1ZW", "Invoice Date": "07-Apr-2025", "Taxable Amount":  48000, "GST Amount":  7680},  # GST mismatch (wrong rate applied)
    {"Invoice No": "INV-005", "Vendor Name": "Wipro Limited",                    "GSTIN": "29AAACW0035A1ZM", "Invoice Date": "08-Apr-2025", "Taxable Amount": 200000, "GST Amount": 36000},  # Matched
    # INV-006 missing in 2A (vendor not filed)
    {"Invoice No": "INV-007", "Vendor Name": "Mahindra & Mahindra Ltd",          "GSTIN": "27AAACM3025E1ZN", "Invoice Date": "12-Apr-2025", "Taxable Amount":  32000, "GST Amount":  5760},  # Matched
    {"Invoice No": "INV-008", "Vendor Name": "HDFC Bank Limited",                "GSTIN": "27AAHCH4118A4ZN", "Invoice Date": "14-Apr-2025", "Taxable Amount":  18000, "GST Amount":  3240},  # Matched
    # INV-009 missing in 2A
    {"Invoice No": "INV-010", "Vendor Name": "Oracle India Pvt Ltd",             "GSTIN": "07AAACO0096F1ZZ", "Invoice Date": "18-Apr-2025", "Taxable Amount":  95000, "GST Amount": 17100},  # Matched
    # INV-011 missing in 2A
    {"Invoice No": "INV-012", "Vendor Name": "Bajaj Electricals Ltd",            "GSTIN": "27AAACB4252E1ZM", "Invoice Date": "22-Apr-2025", "Taxable Amount":  28500, "GST Amount":  5130},  # Matched
    {"Invoice No": "INV-013", "Vendor Name": "Bosch Limited",                    "GSTIN": "29AAACB4174N1Z7", "Invoice Date": "25-Apr-2025", "Taxable Amount":  33000, "GST Amount":  5940},  # Extra in 2A (not in books)
]

pd.DataFrame(books_data).to_excel("purchase_register.xlsx", index=False)
pd.DataFrame(gstr2a_data).to_excel("gstr_2a.xlsx", index=False)
print("Sample files created.")
