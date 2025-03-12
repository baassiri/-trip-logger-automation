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

    # Find the next empty row in the sheet
    last_used_row = 6  # Start from row 7 (after the header)
    
    for row in range(7, ws.max_row + 1):  
        if ws[f"A{row}"].value:
            last_used_row = row

    # Ensure the next entry is always in the next empty row
    new_entry_row = last_used_row + 1

    for client in detected_clients:
        addresses = detected_addresses.get(client, [])
        row_found = None

        # Check if the client already has an entry for the same date
        for row in range(7, last_used_row + 1):
            if ws[f"A{row}"].value == current_date and ws[f"B{row}"].value == client:
                row_found = row
                break

        if row_found is None:
            # Append a new row for this client
            ws[f"A{new_entry_row}"].value = current_date
            ws[f"B{new_entry_row}"].value = client

            for i, address in enumerate(addresses[:5]):  # Up to 5 destinations
                ws.cell(row=new_entry_row, column=5 + i, value=address)

            print(f"🆕 Created new entry for {client} at row {new_entry_row} with addresses: {addresses}")
            new_entry_row += 1  # Move to next available row for next client
        else:
            # Append new addresses to an existing entry
            for address in addresses:
                for col in range(5, 10):  # E to I columns
                    if ws.cell(row=row_found, column=col).value is None:
                        ws.cell(row=row_found, column=col, value=address)
                        print(f"📌 Added {address} to {client} on {current_date} in column {chr(64 + col)}")
                        break
    
    # Save and upload changes
    try:
        wb.save(FILE_PATH)
        wb.close()
        print("✅ Local Excel file updated successfully!")
    except Exception as e:
        print(f"⚠️ Error saving Excel file: {e}")

    upload_to_drive()  # Upload updated file to Google Drive
