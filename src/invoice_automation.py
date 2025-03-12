from openpyxl import load_workbook
from datetime import datetime
import os
from config import FILE_PATH, upload_to_drive

def force_update_trip_log(detected_clients, detected_addresses):
    """Write client names & structured addresses into 'TRIP LOGS' sheet.
    If the same client appears again on the same date, append new addresses into the next available Destination columns.
    Destination columns are E to I (i.e. Destination 1 to Destination 5).
    """
    if not os.path.exists(FILE_PATH):
        print(f"❌ Excel file not found: {FILE_PATH}")
        return

    try:
        wb = load_workbook(FILE_PATH, keep_vba=True)
    except Exception as e:
        print(f"⚠️ Error loading Excel file: {e}")
        return

    if "TRIP LOGS" not in wb.sheetnames:
        print("❌ 'TRIP LOGS' sheet not found in Excel file.")
        wb.close()
        return

    ws = wb["TRIP LOGS"]
    current_date = datetime.now().strftime("%m/%d/%Y")

    # Process each client
    for client in detected_clients:
        addresses = detected_addresses.get(client, [])
        # Try to find an existing row with the same date and client name
        row_found = None
        for row in range(7, ws.max_row + 1):
            if ws[f"A{row}"].value == current_date and ws[f"B{row}"].value == client:
                row_found = row
                break

        if row_found is None:
            # No existing row; create a new row at the bottom.
            new_row = ws.max_row + 1
            ws[f"A{new_row}"].value = current_date
            ws[f"B{new_row}"].value = client
            # Write the addresses into Destination 1-5 (Columns E to I)
            for i, address in enumerate(addresses[:5]):
                ws.cell(row=new_row, column=5 + i, value=address)
            print(f"🔹 New entry created for {client} on {current_date} with addresses: {addresses}")
        else:
            # Row exists; find the first empty destination column (Columns E to I)
            for address in addresses:
                placed = False
                for col in range(5, 10):  # Columns E (5) to I (9)
                    if ws.cell(row=row_found, column=col).value is None:
                        ws.cell(row=row_found, column=col, value=address)
                        placed = True
                        print(f"🔹 Added address for {client} on {current_date} in column {col}")
                        break
                if not placed:
                    print(f"⚠️ All destination columns for {client} on {current_date} are filled. Could not add address: {address}")

    try:
        wb.save(FILE_PATH)
        print("✅ Trip log updated successfully!")
    except Exception as e:
        print(f"⚠️ Error saving Excel file: {e}")
    wb.close()

    # Upload the updated file to Google Drive
    upload_to_drive()
