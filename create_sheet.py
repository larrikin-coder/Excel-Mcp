# from openpyxl import load_workbook,Workbook
# import os
from modules import *

def create_sheet(filepath: str, sheet_name: str):
    """Create a new sheet, create file if it doesn't exist."""
    if not os.path.exists(filepath):
        wb = Workbook()
        wb.create_sheet(title=sheet_name)
        default_sheet = wb["Sheet"]
        wb.remove(default_sheet)
    else:
        wb = load_workbook(filepath)
        wb.create_sheet(title=sheet_name)

    wb.save(filepath)
    return {"message": f"Sheet '{sheet_name}' created in '{filepath}'."}