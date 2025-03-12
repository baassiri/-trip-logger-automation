from openpyxl import load_workbook
from datetime import datetime
import os
from config import FILE_PATH, upload_to_drive

def force_update_trip_log(detected_clients, detected_addresses):
    """
    Write client names & structured addresses into 'TRIP LOGS' sheet.
    If the same client appears again on the same date, append new addresses
    into the next available Destination columns (E to I).
    Ensures entries are written sequentially in order, without skipping rows.
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

    last_used_row = 7  # Start from the first entry row
    
    for row in range(7, ws.max_row + 1):  # Find the last non-empty row
        if ws[f"A{row}"].value:
            last_used_row = row
    
    for client in detected_clients:
        addresses = detected_addresses.get(client, [])
        row_found = None

        # Find existing entry for this client on the same date
        for row in range(7, last_used_row + 1):
            if ws[f"A{row}"].value == current_date and ws[f"B{row}"].value == client:
                row_found = row
                break

        if row_found is None:
            # Insert a new row immediately after the last used row
            new_row = last_used_row + 1
            ws[f"A{new_row}"].value = current_date
            ws[f"B{new_row}"].value = client

            for i, address in enumerate(addresses[:5]):
                ws.cell(row=new_row, column=5 + i, value=address)

            print(f"🆕 Created new entry for {client} with addresses: {addresses}")
            last_used_row += 1  # Update last used row
        else:
            # Append addresses to existing entry
            for address in addresses:
                for col in range(5, 10):  # E to I
                    if ws.cell(row=row_found, column=col).value is None:
                        ws.cell(row=row_found, column=col, value=address)
                        print(f"📌 Added {address} to {client} on {current_date}")
                        break
    
    # Save changes and confirm
    try:
        wb.save(FILE_PATH)
        wb.close()
        print("✅ Local Excel file updated successfully!")
    except Exception as e:
        print(f"⚠️ Error saving Excel file: {e}")

    upload_to_drive()  # Attempt upload to Google Drive