def populate_balance_sheet():
    """Populate the BALANCE SHEET with invoice numbers from INVOICE TRACKER."""
    wb = load_workbook(FILE_PATH, keep_vba=True)
    ws_balance = wb["BALANCE SHEET"]
    ws_invoice_tracker = wb["INVOICE TRACKER"]

    client_name = ws_balance["B4"].value.strip() if ws_balance["B4"].value else ""
    if not client_name:
        print("⚠️ Please enter a client name in B4.")
        return

    start_row = 7
    ref_row = next((row for row in range(7, ws_balance.max_row + 1)
                    if ws_balance[f"A{row}"].value == "__"), None)

    if not ref_row:
        print("⚠️ Reference row '__' not found in Column A.")
        return

    next_row = start_row
    found = False

    for row in range(2, ws_invoice_tracker.max_row + 1):
        if ws_invoice_tracker[f"C{row}"].value == client_name:
            found = True
            if next_row >= ref_row:
                ws_balance.insert_rows(ref_row)
                ref_row += 1
            ws_balance[f"A{next_row}"] = ws_invoice_tracker[f"A{row}"].value
            next_row += 1

    wb.save(FILE_PATH)
    wb.close()
    
    print("✅ Balance Sheet updated!" if found else "⚠️ No matching invoices found.")
