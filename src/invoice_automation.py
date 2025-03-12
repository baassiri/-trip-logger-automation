from openpyxl import load_workbook
from datetime import datetime
import os
from config import FILE_PATH, upload_to_drive

def force_update_trip_log(detected_clients, detected_addresses):
    """
    Write client names & structured addresses into 'TRIP LOGS' sheet.
    - Ensures entries are logged in order from the first available row.
    - If a client is re-logged on the same date, new destinations are added to available columns.
    """
    if not os.path.exists(FILE_PATH):
        print(f"❌ Excel file not found: {FILE_PATH}")
        return

    try:
        wb = load_workbook(FILE_PATH, keep_vba=True)
        ws = wb["TRIP LOGS"]
    except Exception as e:
        print(f"⚠️ Error loading Excel file: {e}")
        return

    current_date = datetime.now().strftime("%m/%d/%Y")
    print(f"📅 Processing trip logs for: {current_date}")

    # Find the first completely empty row starting from row 7
    last_used_row = 7
    while ws[f"A{last_used_row}"].value:
        last_used_row += 1

    for client in detected_clients:
        addresses = detected_addresses.get(client, [])
        row_found = None

        # Look for an existing row for the same client on the same date
        for row in range(7, last_used_row):
            if ws[f"A{row}"].value == current_date and ws[f"B{row}"].value == client:
                row_found = row
                break

        if row_found is None:
            # Insert a new row in order
            ws[f"A{last_used_row}"].value = current_date
            ws[f"B{last_used_row}"].value = client

            for i, address in enumerate(addresses[:5]):  # Max 5 destinations
                ws.cell(row=last_used_row, column=5 + i, value=address)

            print(f"🆕 Created new entry for {client} with addresses: {addresses}")
            last_used_row += 1  # Move to next row for new clients
        else:
            # Append new destinations to existing client row
            for address in addresses:
                for col in range(5, 10):  # Columns E (5) to I (9)
                    if ws.cell(row=row_found, column=col).value is None:
                        ws.cell(row=row_found, column=col, value=address)
                        print(f"📌 Added {address} to {client} on {current_date}")
                        break

    # Save and upload the updated file
    try:
        wb.save(FILE_PATH)
        wb.close()
        print("✅ Local Excel file updated successfully!")
    except Exception as e:
        print(f"⚠️ Error saving Excel file: {e}")

    upload_to_drive()  # Upload updated sheet to Google Drive
