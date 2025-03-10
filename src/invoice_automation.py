# src/invoice_automation.py
import os
from openpyxl import load_workbook
from datetime import datetime
from config import FILE_PATH

def force_update_trip_log(detected_clients, detected_addresses):
    """Write client names & addresses into 'TRIP LOGS' sheet in Excel."""
    if not os.path.exists(FILE_PATH):
        print(f"❌ Excel file not found: {FILE_PATH}")
        return

    wb = load_workbook(FILE_PATH, keep_vba=True)
    ws = wb["TRIP LOGS"]

    current_date = datetime.now().strftime("%m/%d/%Y")

    # Find first empty row in column B (Client)
    empty_row = 7
    while ws[f"B{empty_row}"].value:
        empty_row += 1

    for client in detected_clients:
        address = detected_addresses.get(client, "Unknown Address")
        print(f"🔹 Logging {client} with address '{address}' at row {empty_row}")

        ws[f"A{empty_row}"].value = current_date
        ws[f"B{empty_row}"].value = client
        ws[f"E{empty_row}"].value = address

        empty_row += 1

    wb.save(FILE_PATH)
    wb.close()
    print("✅ Trip log updated!")
