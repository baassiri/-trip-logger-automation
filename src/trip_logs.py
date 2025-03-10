# src/trip_logs.py
from openpyxl import load_workbook
from datetime import datetime
from config import FILE_PATH

def clear_trip_logs():
    """Example function to clear trip logs in Excel."""
    wb = load_workbook(FILE_PATH, keep_vba=True)
    ws = wb["TRIP LOGS"]

    last_row = ws.max_row
    for row in range(7, last_row + 1):
        ws.delete_rows(row)

    wb.save(FILE_PATH)
    wb.close()
    print("✅ Trip logs cleared.")
