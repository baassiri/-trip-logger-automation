from openpyxl import load_workbook
from config import FILE_PATH

def clear_trip_logs():
    """
    Clears the trip logs in the 'TRIP LOGS' sheet of the Excel file.
    It deletes all rows starting from row 7.
    """
    wb = load_workbook(FILE_PATH, keep_vba=True)
    ws = wb["TRIP LOGS"]

    last_row = ws.max_row
    if last_row >= 7:
        # Delete rows starting at row 7 and delete all rows until the end
        ws.delete_rows(7, last_row - 7 + 1)

    wb.save(FILE_PATH)
    wb.close()
    print("✅ Trip logs cleared.")
