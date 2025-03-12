from openpyxl import load_workbook
from datetime import datetime
import os
from config import FILE_PATH, upload_to_drive


def force_update_trip_log(detected_clients, detected_addresses):
    """Write client names & multiple structured addresses into 'TRIP LOGS' sheet in Excel and upload changes to Google Drive."""
    
    if not FILE_PATH or not os.path.exists(FILE_PATH):
        print(f"❌ Excel file not found: {FILE_PATH}")
        return
    
    try:
        wb = load_workbook(FILE_PATH, keep_vba=True)
    except Exception as e:
        print(f"⚠️ Error loading Excel file: {e}")
        return
    
    # Ensure 'TRIP LOGS' sheet exists
    if "TRIP LOGS" not in wb.sheetnames:
        print("❌ 'TRIP LOGS' sheet not found in Excel file.")
        wb.close()
        return

    ws = wb["TRIP LOGS"]
    current_date = datetime.now().strftime("%m/%d/%Y")

    # Find first empty row in column B (Client)
    empty_row = 7
    while ws[f"B{empty_row}"].value:
        empty_row += 1

    for client in detected_clients:
        addresses = detected_addresses.get(client, [])

        print(f"🔹 Logging {client} with addresses: {addresses} at row {empty_row}")

        ws[f"A{empty_row}"].value = current_date  # Log the date
        ws[f"B{empty_row}"].value = client  # Log the client name

        # Log up to 5 structured addresses in columns E, F, G, H, I
        for i, address in enumerate(addresses[:5]):  # Maximum of 5 destinations
            ws.cell(row=empty_row, column=5 + i, value=address)

        empty_row += 1  # Move to the next row for new entries

    try:
        wb.save(FILE_PATH)
        print("✅ Trip log updated successfully!")
    except Exception as e:
        print(f"⚠️ Error saving Excel file: {e}")
    
    wb.close()

    # Upload updated file back to Google Drive
    upload_to_drive()
