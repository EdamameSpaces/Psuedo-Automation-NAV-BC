
import pandas as pd
import re
from datetime import datetime
import os

# NAV Column Headers
NAV_COLUMNS = [
    "Posting Date", "Document Type", "Document No.", "Account Type", "Account No.",
    "Description", "Gen. Posting Group", "Gen. Bus. Posting Group", "Gen. Prod. Posting Group",
    "Amount", "Bal. Account Type", "Bal. Account No."
]

# Walmart G/L Account Mapping
GL_ACCOUNT_MAP_WALMART = {
    "0100": "482100",
    "0022": "486000",
    "0024": "486000",
    "0780": "636300",
    "0059": "488000",
    "0057": "482100",
    "0775": "825000",
    "0025": "486000",
    "0088": "485300",
    "0762": "825000",
    "0130": "482100",
    "0054": "482100",
    "0087": "825000"
}

# Abbreviated Descriptions
ABBREV_DESC_MAP = {
    "0100": "UNSEAL/QUANTITY ALLOWANCE",
    "0022": "MERCHANDISE SHORTAGE",
    "0024": "CARTON SHORTAGE FREIGHT",
    "0780": "TRANSPORT BILLING",
    "0059": "DEFECTIVE ALLOWANCE",
    "0057": "QUANTITY DISC ALLOWANCE",
    "0775": "MARKDOWN BILLING",
    "0025": "POD/NO MERCHANDISE SHORTAGE",
    "0088": "MERCHANDISE RETURNS",
    "0762": "COMPLIANCE BILLING",
    "0130": "SUBSTITUTION OVERCHARGE",
    "0054": "WAREHOUSE ALLOWANCE",
    "0087": "OTHER"
}

# Codes to Sum
SUM_CODES = ["0100", "0057", "0059"]

def extract_code(description):
    if pd.isna(description):
        return None
    match = re.search(r"\[(\d{4})\]", str(description))
    return match.group(1) if match else None

def clean_amount(val):
    s = str(val).replace(",", "").replace("(", "-").replace(")", "").strip()
    return pd.to_numeric(s, errors="coerce")

def process_walmart_file(df, payment_number, posting_date):
    df["Deduction Code"] = df["DEDUCTION CODE"].apply(extract_code)
    df["Amount"] = df["Amount Paid($)"].apply(clean_amount)

    rows = []

    for code in SUM_CODES:
        subset = df[df["Deduction Code"] == code]
        total = subset["Amount"].sum()
        if total != 0:
            desc = f"PMT {payment_number} {ABBREV_DESC_MAP.get(code, code)}"
            for amt in [total, -total]:
                rows.append([
                    posting_date, " ", " ", "Customer", "8501", desc,
                    " ", " ", " ", amt, "G/L Account", GL_ACCOUNT_MAP_WALMART[code]
                ])

    other_df = df[~df["Deduction Code"].isin(SUM_CODES)]

    for _, row in other_df.iterrows():
        code = row["Deduction Code"]
        gl_code = GL_ACCOUNT_MAP_WALMART.get(code)
        if not gl_code or pd.isna(row["Amount"]):
            continue
        abbrev_desc = ABBREV_DESC_MAP.get(code, code)
        desc = f"PMT {payment_number} {row['Invoice Number']} {abbrev_desc}"
        for amt in [row["Amount"], -row["Amount"]]:
            rows.append([
                posting_date, " ", " ", "Customer", "8501", desc,
                " ", " ", " ", amt, "G/L Account", gl_code
            ])

    return pd.DataFrame(rows, columns=NAV_COLUMNS)

if __name__ == "__main__":
    file_path = input("Enter path to Walmart remittance .xlsx file: ").strip()
    df = pd.read_excel(file_path)

    file_name = os.path.basename(file_path)
    check_number_match = re.search(r"Check_(\d+)", file_name)
    check_number = check_number_match.group(1) if check_number_match else "WMT001"

    df["Date Paid"] = pd.to_datetime(df["Date Paid"], errors='coerce')
    posting_date = df["Date Paid"].max().strftime('%m/%d/%Y')
    payment_amount = abs(df["Amount Paid($)"].apply(clean_amount).sum())

    result_df = process_walmart_file(df, check_number, posting_date)
    output_path = f"{check_number}_{payment_amount:.2f}.xlsx"
    result_df.to_excel(output_path, index=False)
    print(f"NAV export saved to: {output_path}")
