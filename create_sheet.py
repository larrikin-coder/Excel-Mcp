from modules import *

def create_sheet(filepath: str, sheet_name: str):
    """Create a new sheet, create file if it doesn't exist."""

    if not os.path.exists(filepath):
        wb = Workbook()
        
        # Remove default sheet before creating the new sheet
        default_sheet = wb["Sheet"]
        wb.remove(default_sheet)
        wb.save(filepath)  # Save after removing default sheet
        
        # Now create the requested new sheet
        wb.create_sheet(title=sheet_name)
        wb.save(filepath)  # Save again after adding new sheet

    else:
        wb = load_workbook(filepath)
        wb.create_sheet(title=sheet_name)
        wb.save(filepath)

    return {"message": f"Sheet '{sheet_name}' created in '{filepath}'."}
