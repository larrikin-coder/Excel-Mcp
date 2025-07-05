# from openpyxl import load_workbook
from modules import *

def write_cell(filepath: str, sheet_name: str, cell: str, value: str):
    """Write a value to a specific cell. Create sheet if missing. Handle spaces, cases safely."""
    wb = load_workbook(filepath)

    clean_sheet_names = {s.strip().lower(): s for s in wb.sheetnames}
    requested_clean = sheet_name.strip().lower()

    if requested_clean in clean_sheet_names:
        ws = wb[clean_sheet_names[requested_clean]]
    else:
        ws = wb.create_sheet(title=sheet_name)

    ws[cell] = value
    wb.save(filepath)
    return {"message": f"Value '{value}' written to {sheet_name}:{cell}"}