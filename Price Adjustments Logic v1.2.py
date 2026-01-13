
from openpyxl import Workbook

def process_price_adjustments_from_prices(price_data, invoice_number):
    """
    price_data: list of tuples (item_no, PO_price, Invoice_price, qty)
    invoice_number: str
    """
    wb = Workbook()

    # Sheet 1: Audit Trail
    ws = wb.active
    ws.title = "Sheet1"
    headers = [
        'Type', 'No.', 'Description', 'Location Code', 'Quantity', 'Unit of Measure Code',
        'Unit Price Excl. Tax', 'Tax Group Code', 'Line Amount Excl. Tax', 'Amount Including Tax',
        'Line Discount %', 'Qty. to Assign', 'Qty. Assigned'
    ]
    ws.append(headers)

    adjustments = []
    for item_no, po_price, inv_price, qty in price_data:
        overcharge = round(inv_price - po_price, 2)
        if overcharge <= 0:
            continue

        total_line = qty * overcharge
        adjustments.append((invoice_number, item_no, qty, overcharge))

        unit_price_fmt = f"{overcharge:.2f}"
        line_amount_fmt = f"{total_line:.2f}"
        description = f"({qty}) {item_no} @ ${unit_price_fmt} EA {invoice_number}"

        ws.append([
            "G/L Account", 482100, "PRICE ADJUSTMENTS", "W01", qty, "EA",
            unit_price_fmt, "NONTAXABLE", line_amount_fmt, line_amount_fmt,
            " ", 0, " "
        ])
        ws.append([
            " ", " ", description, " ", " ", " ",
            " ", " ", " ", 0,
            " ", 0, " "
        ])

    ws.append([" ", " ", f"PRICE ADJUSTMENT {invoice_number}", " ", " ", " ", " ", " ", " ", 0, " ", 0, " "])

    # Sheet 2: Backup Calculations
    ws_math = wb.create_sheet(title="Backup Calculations")
    ws_math.append([
        "Item Number", "Qty", "PO Unit Price", "Invoiced Unit Price",
        "Overcharge per Unit", "Total Overcharge"
    ])

    for idx, (item_no, po_price, inv_price, qty) in enumerate(price_data, start=2):
        ws_math.append([
            item_no,
            qty,
            po_price,
            inv_price,
            f"=D{idx}-C{idx}",
            f"=B{idx}*E{idx}"
        ])

    output_path = f"/mnt/data/{invoice_number} Price Adjustment.xlsx"
    wb.save(output_path)
    return output_path
