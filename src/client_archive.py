import os
from openpyxl import load_workbook
from config import FILE_PATH

def move_to_archive():
    """
    Move clients marked as 'archive' from CLIENT DATA to CLIENT ARCHIVE.
    VBA Equivalent: MoveToArchive()
    """
    if not os.path.exists(FILE_PATH):
        print(f"❌ Error: Excel file not found at {FILE_PATH}")
        return

    wb = load_workbook(FILE_PATH, keep_vba=True)
    ws_client_data = wb["CLIENT DATA"]
    ws_client_archive = wb["CLIENT ARCHIVE"]

    # Find last row in CLIENT DATA (Column A)
    last_row_data = ws_client_data.max_row

    # Loop backward from last row down to row 2
    for i in range(last_row_data, 1, -1):
        cell_value = ws_client_data.cell(row=i, column=21).value  # 'U' = 21st column
        if cell_value and str(cell_value).strip().lower() == "archive":
            # Next available row in CLIENT ARCHIVE
            last_row_archive = ws_client_archive.max_row + 1
            if last_row_archive < 6:
                last_row_archive = 6

            # Copy row (A:U => columns 1..21)
            row_data = []
            for col in range(1, 22):
                row_data.append(ws_client_data.cell(row=i, column=col).value)

            # Paste into CLIENT ARCHIVE
            for col in range(1, 22):
                ws_client_archive.cell(row=last_row_archive, column=col, value=row_data[col - 1])

            # Verify successful copy
            if ws_client_archive.cell(row=last_row_archive, column=1).value == ws_client_data.cell(row=i, column=1).value:
                # Delete the row from CLIENT DATA
                ws_client_data.delete_rows(i)
            else:
                print(f"⚠️ Error: Row {i} could not be copied to CLIENT ARCHIVE. Deletion skipped.")

    wb.save(FILE_PATH)
    wb.close()
    print("✅ Clients marked as 'archive' have been moved successfully!")

def restore_from_archive():
    """
    Move clients marked as 'active' from CLIENT ARCHIVE back to CLIENT DATA.
    VBA Equivalent: RestoreFromArchive()
    """
    if not os.path.exists(FILE_PATH):
        print(f"❌ Error: Excel file not found at {FILE_PATH}")
        return

    wb = load_workbook(FILE_PATH, keep_vba=True)
    ws_client_data = wb["CLIENT DATA"]
    ws_client_archive = wb["CLIENT ARCHIVE"]

    # Find last row in CLIENT ARCHIVE (Column A)
    last_row_archive = ws_client_archive.max_row

    # Loop backward from last_row_archive down to row 6
    for i in range(last_row_archive, 5, -1):
        cell_value = ws_client_archive.cell(row=i, column=21).value  # 'U' = 21st column
        if cell_value and str(cell_value).strip().lower() == "active":
            # Next available row in CLIENT DATA
            last_row_data = ws_client_data.max_row + 1
            if last_row_data < 2:
                last_row_data = 2

            # Copy row (A:U => columns 1..21)
            row_data = []
            for col in range(1, 22):
                row_data.append(ws_client_archive.cell(row=i, column=col).value)

            # Paste into CLIENT DATA
            for col in range(1, 22):
                ws_client_data.cell(row=last_row_data, column=col, value=row_data[col - 1])

            # Verify successful copy
            if ws_client_data.cell(row=last_row_data, column=1).value == ws_client_archive.cell(row=i, column=1).value:
                # Delete the row from CLIENT ARCHIVE
                ws_client_archive.delete_rows(i)
            else:
                print(f"⚠️ Error: Row {i} could not be copied back to CLIENT DATA. Deletion skipped.")

    wb.save(FILE_PATH)
    wb.close()
    
