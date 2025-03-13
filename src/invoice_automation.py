from openpyxl import load_workbook
from datetime import datetime
import os
from config import FILE_PATH, upload_to_drive

print("✅ Script started!")

def force_update_trip_log(detected_clients, detected_addresses):
    """Append trip logs after the last used row in Column A (starting from row 7)."""

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

    # 1️⃣ Find the last used row in Column A (>= row 7).
    last_row = 7
    for row in range(7, ws.max_row + 1):
        if ws.cell(row=row, column=1).value:  # Column A is column=1
            last_row = row

    # 2️⃣ We'll insert new entries starting at last_row + 1
    new_row = last_row + 1

    # Process each client
    for client in detected_clients:
        addresses = detected_addresses.get(client, [])
        row_found = None

        # 3️⃣ Check if this client/date already exists in rows 7..last_row
        for row in range(7, last_row + 1):
            if ws.cell(row=row, column=1).value == current_date and ws.cell(row=row, column=2).value == client:
                row_found = row
                break

        if row_found is None:
            # 4️⃣ Insert a new entry at new_row
            ws.cell(row=new_row, column=1, value=current_date)  # Date in Col A
            ws.cell(row=new_row, column=2, value=client)        # Client in Col B

            for i, address in enumerate(addresses[:5]):  # Up to 5 destinations
                ws.cell(row=new_row, column=5 + i, value=address)

            print(f"🆕 Created new entry for {client} at row {new_row} with addresses: {addresses}")
            new_row += 1  # Next new entry goes one row down
        else:
            # 5️⃣ If found, append addresses in the next available columns (E..I)
            for address in addresses:
                for col in range(5, 10):  # E=5..I=9
                    if not ws.cell(row=row_found, column=col).value:
                        ws.cell(row=row_found, column=col, value=address)
                        print(f"📌 Added {address} to {client} at row {row_found}, column {col}")
                        break

    # (Optional) Print the last few rows for verification
    print("🔍 Verifying last few rows before saving:")
    for row in range(max(7, new_row - 5), new_row):
        row_values = [ws.cell(row=row, column=col).value for col in range(1, 11)]
        print(row_values)

    # 6️⃣ Save and upload
    try:
        wb.save(FILE_PATH)
        wb.close()
        print("✅ Local Excel file updated successfully!")
    except Exception as e:
        print(f"⚠️ Error saving Excel file: {e}")

    upload_to_drive()
