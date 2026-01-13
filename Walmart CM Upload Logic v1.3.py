import pandas as pd
import re

# Mapping descriptions to account numbers
CHARGEBACK_ACCOUNT_MAP = {
    "MULTI INVOICE PRICE DIFFERENCE": ("482100", "PRICE ADJUSTMENTS"),
    "UNSEAL/QUANTITY ALLOWANCE": ("482100", "PRICE ADJUSTMENTS"),
    "QUANTITY DISC ALLOWANCE": ("482100", "PRICE ADJUSTMENTS"),
    "DEFECTIVE ALLOWANCE": ("488000", "DEFECTIVE ALLOWANCES"),
    "MARKDOWN BILLING": ("825000", "VENDOR VIOLATIONS"),
    "WAREHOUSE ALLOWANCE": ("482100", "PRICE ADJUSTMENTS"),
    "MERCHANDISE SHORTAGE": ("486000", "SHORTAGES"),
    "MECHANDISE SHORTAGE": ("486000", "SHORTAGES"),
    "CARTON SHORTAGE FREIGHT": ("486000", "SHORTAGES"),
    "POD/NO MERCHANDISE SHORTAGE": ("486000", "SHORTAGES"),
}

# Abbreviations for long chargeback descriptions
DESCRIPTION_ABBREVIATIONS = {
    "UNSEAL/QUANTITY ALLOWANCE": "UNSEAL/QTY ALLOW",
    "QUANTITY DISC ALLOWANCE": "QTY DISC ALLOW",   
    "DEFECTIVE ALLOWANCE": "DEFECTIVE ALLOW",
    "MARKDOWN BILLING": "MARKDOWN BILL",
    "WAREHOUSE ALLOWANCE": "WAREHOUSE ALLOW",
    "MERCHANDISE SHORTAGE": "MERCHANDISE SHORT",
    "MECHANDISE SHORTAGE": "MERCHANDISE SHORT",
    "CARTON SHORTAGE FREIGHT": "FREIGHT SHORT",
    "POD/NO MERCHANDISE SHORTAGE": "POD SHORTAGE",
}

def abbreviate_and_truncate(description, max_length=35):
    # Apply abbreviation if applicable
    for long_desc, short_desc in DESCRIPTION_ABBREVIATIONS.items():
        if long_desc in description.upper():
            description = description.upper().replace(long_desc, short_desc)
    return description[:max_length]

def read_txt_file(file_path):
    records = []
    with open(file_path, 'r') as file:
        for line in file:
            parts = re.split(r'\s{2,}', line.strip())
            if len(parts) == 3:
                customer_no, chargeback_no, amount_str = parts
                amount = float(re.sub(r'[^\d.-]', '', amount_str))
                records.append({
                    "CustomerNumber": customer_no,
                    "ChargebackNumber": chargeback_no,
                    "Amount": amount
                })
    return records

def generate_ascr_numbers(start_ascr, count):
    prefix, start_num = start_ascr.split('-')
    return [f"{prefix}-{int(start_num) + i:06d}" for i in range(count)]

def find_account_info(description):
    for key_phrase, (gl_no, desc) in CHARGEBACK_ACCOUNT_MAP.items():
        if key_phrase in description.upper():
            return gl_no, desc
    return "UNKNOWN", "UNKNOWN"

def populate_sales_header(records, ascr_numbers, descriptions):
    header_data = []
    for record, ascr, desc in zip(records, ascr_numbers, descriptions):
        truncated_desc = abbreviate_and_truncate(desc)
        header_data.append({
            "Document Type": "Credit Memo",
            "No.": ascr,
            "Sell-to Customer No.": record["CustomerNumber"],
            "Bill-to Customer No.": record["CustomerNumber"],
            "Posting Description": f"Credit Memo {ascr}",
            "Location Code": "RFID",
            "Applies-to Doc. Type": "",  # Removed "Payment"
            "Applies-to Doc. No.": record["ChargebackNumber"],
            "External Document No.": truncated_desc
        })
    return pd.DataFrame(header_data)

def populate_sales_line(records, ascr_numbers, descriptions):
    line_data = []
    for record, ascr, desc in zip(records, ascr_numbers, descriptions):
        gl_no, gl_desc = find_account_info(desc)
        truncated_desc = abbreviate_and_truncate(desc)
        line_data.append({
            "Document Type": "Credit Memo",
            "Document No.": ascr,
            "Line No.": 10000,
            "Type": "G/L Account",
            "No.": gl_no,
            "Location Code": "RFID",
            "Description": gl_desc,
            "Quantity": 1,
            "Unit Price": record["Amount"],
            "Amount": record["Amount"],
            "Tax Group Code": "NONTAXABLE"
        })
        line_data.append({
            "Document Type": "Credit Memo",
            "Document No.": ascr,
            "Line No.": 20000,
            "Type": "G/L Account",
            "No.": gl_no,
            "Location Code": "RFID",
            "Description": truncated_desc,
            "Quantity": 0,
            "Unit Price": 0,
            "Amount": 0,
            "Tax Group Code": "NONTAXABLE"
        })
    return pd.DataFrame(line_data)

def generate_credit_memo_excel(txt_file_path, start_ascr, descriptions, output_path):
    records = read_txt_file(txt_file_path)
    if len(records) != len(descriptions):
        raise ValueError("Mismatch between record count and description count")
    ascr_numbers = generate_ascr_numbers(start_ascr, len(records))
    header_df = populate_sales_header(records, ascr_numbers, descriptions)
    line_df = populate_sales_line(records, ascr_numbers, descriptions)
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        header_df.to_excel(writer, sheet_name='Sales Header', index=False)
        line_df.to_excel(writer, sheet_name='Sales Line', index=False)
    print(f"Walmart Credit Memo Excel generated: {output_path}")
