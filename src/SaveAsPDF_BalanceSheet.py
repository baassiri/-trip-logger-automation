from win32com.client import Dispatch

def save_as_pdf(sheet_name, output_path):
    """Save a given Excel sheet as a PDF."""
    excel = Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(FILE_PATH)
    ws = wb.Sheets(sheet_name)

    ws.ExportAsFixedFormat(0, output_path)
    wb.Close(SaveChanges=False)
    excel.Quit()
    
    print(f"✅ PDF saved at: {output_path}")
