def reset_balance_sheet():
    """Clear invoice numbers from BALANCE SHEET."""
    wb = load_workbook(FILE_PATH, keep_vba=True)
    ws_balance = wb["BALANCE SHEET"]

    start_row = 7
    ref_row = next((row for row in range(start_row, ws_balance.max_row + 1)
                    if ws_balance[f"A{row}"].value == "__"), None)

    if not ref_row:
        print("⚠️ Reference row '__' not found.")
        return

    ws_balance.delete_rows(start_row, ref_row - start_row)
    wb.save(FILE_PATH)
    wb.close()
    
    print("✅ Balance Sheet reset.")
