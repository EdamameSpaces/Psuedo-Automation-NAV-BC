
import pandas as pd
from datetime import datetime

# Constants
NAV_COLUMNS = [
    "Posting Date", "Document Type", "Document No.", "Account Type", "Account No.",
    "Description", "Gen. Posting Group", "Gen. Bus. Posting Group", "Gen. Prod. Posting Group",
    "Amount", "Bal. Account Type", "Bal. Account No."
]

ABBREV_MAP = {
    "PO on-time accuracy": "PO Accuracy",
    "Prep-Bagging": "Prep-Bagging",
    "Shortage Claim for Invoice": "Shortage Claim",
    "Missed Adjustment Claim for Invoice": "Missed Adjustment Claim",
    "Ship In Own Container": "Ship in Own Container",
    "Price Claim for Invoice": "Price Claim",
    "PROVISION_FOR_RECEIVABLE": "Provision for Receivable",
    "Damage Allowance": "Defective Allowance",
    "Defective": "Defective Allowance",
    "Co-op": "Co-op",
    "Quantity/Bulk Buy Allowance": "Quantity/Bulk Allowance",
    "Bulk Buy Allowance": "Quantity/Bulk Allowance"
    "Incorrect Quantity": "Incorrect Quantity"
}

GL_ACCOUNT_MAP_NORMALIZED = {
    "po on-time accuracy": "825000",
    "prep-bagging": "825000",
    "shortage claim for invoice": "486000",
    "missed adjustment claim for invoice": "486000",
    "ship in own container": "825000",
    "price claim for invoice": "482100",
    "provision_for_receivable": "109500",
    "damage allowance": "488000",
    "defective": "488000",
    "co-op": "226000",
    "quantity/bulk buy allowance": "482100",
    "bulk buy allowance": "482100"
    "Incorrect Quantity": "485300"
}

# Utility functions
def clean_amount(value):
    s = str(value).replace("*", "").replace(",", "").strip()
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    return pd.to_numeric(s, errors="coerce")

def calculate_chargeback_amount(entry):
    paid = clean_amount(entry.get("Amount Paid", 0))
    remaining = clean_amount(entry.get("Amount Remaining", 0))
    return paid + remaining

def extract_base_description(description):
    desc_lower = description.lower()
    if "reverse for" in desc_lower or "reversal for" in desc_lower:
        return None  # Explicitly exclude reversal lines
    if "co-op" in desc_lower:
        return "Co-op"
    if "prep - bagging" in desc_lower or "prep-bagging" in desc_lower:
        return "Prep-Bagging"
    if "shortage claim for invoice" in desc_lower:
        return "Shortage Claim for Invoice"
    if "missed adjustment claim for invoice" in desc_lower:
        return "Missed Adjustment Claim for Invoice"
    if "ship in own container" in desc_lower:
        return "Ship In Own Container"
    if "po on-time accuracy" in desc_lower:
        return "PO on-time accuracy"
    if "provision_for_receivable" in desc_lower:
        return "PROVISION_FOR_RECEIVABLE"
    if "damage allowance" in desc_lower:
        return "Damage Allowance"
    if "price claim for invoice" in desc_lower:
        return "Price Claim"
    if "quantity/bulk buy allowance" in desc_lower:
        return "Quantity/Bulk Buy Allowance"
    if "bulk buy allowance" in desc_lower:
        return "Bulk Buy Allowance"
    if " - " in description:
        return description.split(" - ")[0].strip()
    return description.split(",")[0].strip()

def extract_base_description_normalized(description):
    base = extract_base_description(description)
    return base.strip().lower() if base else None

def generate_description(payment_number, invoice_number, full_description):
    base_desc = extract_base_description(full_description)
    if not base_desc:
        return None
    abbrev = ABBREV_MAP.get(base_desc, base_desc)
    return f"PMT {payment_number} {invoice_number} {abbrev}"

def process_chargebacks(data, payment_number, payment_amount, posting_date):
    rows = []
    for entry in data:
        if "*" in str(entry.get("Amount Paid", "")):
            continue
        amount = calculate_chargeback_amount(entry)
        if amount == 0 or pd.isna(amount):
            continue
        base_desc = extract_base_description_normalized(entry["Description"])
        if not base_desc:
            continue
        gl_account = GL_ACCOUNT_MAP_NORMALIZED.get(base_desc, None)
        if not gl_account:
            continue
        abbrev = ABBREV_MAP.get(base_desc.title(), base_desc.title())
        desc = f"PMT {payment_number} {entry['Invoice Number']} {abbrev}"
        for amt in [amount, -amount]:
            rows.append([
                posting_date, " ", " ", "Customer", "1287", desc,
                " ", " ", " ", amt, "G/L Account", gl_account
            ])
    df = pd.DataFrame(rows, columns=NAV_COLUMNS)
    return df


# === PATCHED EXPORT FUNCTION (v3.1.4) ===
def export_chargebacks_to_excel(df, payment_number, payment_amount, export_dir="/mnt/data"):
    """
    Applies final formatting and saves to Excel using desired filename and date format.
    """
    # Ensure date format is mm/dd/yyyy
    df["Posting Date"] = pd.to_datetime(df["Posting Date"]).dt.strftime("%m/%d/%Y")

    # Create filename
    filename = f"{payment_number}_{payment_amount:.2f}.xlsx"
    filepath = f"{export_dir}/{filename}"

    # Save file
    df.to_excel(filepath, index=False)
    return filepath
