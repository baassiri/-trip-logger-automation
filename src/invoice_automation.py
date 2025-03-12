import os
from datetime import datetime
from openpyxl import load_workbook
from config import FILE_PATH, upload_to_drive

def check_local_file():
    """Check if the file was updated before uploading."""
    if os.path.exists(FILE_PATH):
        print(f"📂 File exists at: {FILE_PATH}")
        print(f"📏 File size before upload: {os.path.getsize(FILE_PATH)} bytes")
    else:
        print("❌ Local file not found! Google Drive upload will fail.")

def force_update_trip_log(detected_clients, detected_addresses):
    """Updates the Excel sheet and logs new addresses per client per date."""
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

    for client in detected_clients:
        addresses = detected_addresses.get(client, [])
        row_found = None

        # Find existing entry for this client on the same date
        for row in range(7, ws.max_row + 1):
            if ws[f"A{row}"].value == current_date and ws[f"B{row}"].value == client:
                row_found = row
                break

        if row_found is None:
            # Add a new row for the client
            new_row = ws.max_row + 1
            ws[f"A{new_row}"].value = current_date
            ws[f"B{new_row}"].value = client

            for i, address in enumerate(addresses[:5]):
                ws.cell(row=new_row, column=5 + i, value=address)

            print(f"🆕 Created new entry for {client} with addresses: {addresses}")

        else:
            # Append addresses to existing entry
            for address in addresses:
                for col in range(5, 10):  # Columns E to I
                    if ws.cell(row=row_found, column=col).value is None:
                        ws.cell(row=row_found, column=col, value=address)
                        print(f"📌 Added {address} to {client} on {current_date}")
                        break
                else:
                    print(f"⚠️ No available columns to add new address for {client} on {current_date}")

    # Save changes and confirm
    try:
        wb.save(FILE_PATH)
        wb.close()
        print("✅ Local Excel file updated successfully!")
        check_local_file()  # Verify update before upload
    except Exception as e:
        print(f"⚠️ Error saving Excel file: {e}")

    # Attempt to upload the updated file to Google Drive
    try:
        print("📤 Uploading updated XLSM to Google Drive...")
        upload_to_drive()
        print("✅ Upload successful!")
    except Exception as e:
        print(f"❌ Upload failed: {e}")
