from modules import Workbook, load_workbook, os, file_to_base64

def create_sheet(filepath: str, sheet_name: str):
    """Create a new sheet, create file if it doesn't exist, return updated file in base64."""

    if not os.path.exists(filepath):
        wb = Workbook()
        default_sheet = wb["Sheet"]
        wb.remove(default_sheet)
        wb.create_sheet(title=sheet_name)
        wb.save(filepath)
    else:
        wb = load_workbook(filepath)
        wb.create_sheet(title=sheet_name)
        wb.save(filepath)

    return {
        "message": f"Sheet '{sheet_name}' created in '{filepath}'.",
        "file": file_to_base64(filepath)
    }
