# Excel MCP Agent

This project is an **Excel MCP (Multi-Command Processor) Agent** that allows you to edit and modify Excel files smartly using commands. It provides a set of Python scripts and a possible frontend/server interface to automate and streamline Excel file manipulation.

## Features
- Command-based editing and modification of Excel files
- Smart automation for common Excel tasks
- Modular design for extensibility
- Example scripts for creating, writing, and processing Excel files

## Project Structure
- `create_sheet.py` – Script to create new Excel sheets
- `write_cell.py` – Script to write data to specific cells
- `modules.py` – Contains reusable modules for Excel operations
- `mcp_server.py` – (Optional) Server component for handling commands (possibly via API or frontend)
- `frontend.py` – (Optional) Frontend interface for user interaction
- `multiplication_tables.xlsx`, `output.xlsx`, `uploaded_file.xlsx` – Example Excel files

## Requirements
- Python 3.7+
- [openpyxl](https://pypi.org/project/openpyxl/) (for Excel file manipulation)

Install dependencies with:
```bash
pip install openpyxl
```

## Usage
You can run the scripts directly from the command line. For example:

```bash
python create_sheet.py
python write_cell.py
```

If using the server or frontend, run:
```bash
python mcp_server.py
python frontend.py
```

## Extending
Add new command modules in `modules.py` to support more Excel operations.

## License
MIT License 