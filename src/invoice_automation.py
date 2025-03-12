from openpyxl import load_workbook
from datetime import datetime
import os
from config import FILE_PATH, upload_to_drive

def force_update_trip_log(detected_clients, detected_addresses):
    """
    Write client names & multiple structured addresses into the first 21 columns (A..U) 
    of the 'TRIP LOGS' sheet in Excel, then upload to Google Drive.
    """

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

    # For each client, log up to 5 addresses in columns E..I, but we allow up to 21 columns total
    for client in detected_clients:
        addresses = detected_addresses.get(client, [])

        print(f"🔹 Logging {client} with addresses: {addresses} at row {empty_row}")

        # A..U => 21 columns total
        # Example usage:
        #   A => Date
        #   B => Client
        #   C..D => (optional columns, e.g., Base, Home)
        #   E..I => Up to 5 addresses
        #   J..U => Additional columns if needed
        ws[f"A{empty_row}"].value = current_date  # Date
        ws[f"B{empty_row}"].value = client        # Client

        # Write up to 5 addresses in columns E..I (5 columns)
        for i, address in enumerate(addresses[:5]):  # Maximum 5 destinations
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
