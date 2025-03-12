from openpyxl import load_workbook
from datetime import datetime
import os
from config import FILE_PATH, upload_to_drive

def force_update_trip_log(detected_clients, detected_addresses):
    """
    Append new client trip logs in order.
    If a client appears again on the same date, append new addresses under the next available Destination column.
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

    # Find the last used row dynamically
    last_used_row = max(7, ws.max_row)  # Ensure we start from at least row 7

    for client in detected_clients:
        addresses = detected_addresses.get(client, [])
        row_found = None

        # Check if the client already exists on the same date
        for row in range(7, last_used_row + 1):
            if ws[f"A{row}"].value == current_date and ws[f"B{row}"].value == client:
                row_found = row
                break

        if row_found is None:
            # Append a new entry at the next available row
            new_row = last_used_row + 1
            ws[f"A{new_row}"].value = current_date
            ws[f"B{new_row}"].value = client

            for i, address in enumerate(addresses[:5]):
                ws.cell(row=new_row, column=5 + i, value=address)

            print(f"🆕 Created new entry for {client} with addresses: {addresses}")
            last_used_row += 1  # Move to next row
        else:
            # Append addresses to the next available column
            for address in addresses:
                for col in range(5, 10):  # E to I
                    if ws.cell(row=row_found, column=col).value is None:
                        ws.cell(row=row_found, column=col, value=address)
                        print(f"📌 Added {address} to {client} on {current_date}")
                        break

    # Save changes
    try:
        wb.save(FILE_PATH)
        wb.close()
        print("✅ Local Excel file updated successfully!")
    except Exception as e:
        print(f"⚠️ Error saving Excel file: {e}")

    # Upload updated file to Google Drive
    upload_to_drive()
