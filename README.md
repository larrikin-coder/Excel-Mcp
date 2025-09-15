# Excel MCP Agent

[![GitHub Repo](https://img.shields.io/badge/GitHub-Excel--Mcp-blue?logo=github)](https://github.com/larrikin-coder/Excel-Mcp)

This project is an **Excel MCP (Multi-Command Processor) Agent** that allows you to edit and modify Excel files smartly using commands. It provides a set of Python scripts and a possible frontend/server interface to automate and streamline Excel file manipulation.

## Features
- Command-based editing and modification of Excel files
- Smart automation for common Excel tasks
- Modular design for extensibility
- Example scripts for creating, writing, and processing Excel files

## Getting Started

Follow these steps to run the project on your own PC:

### 1. Clone the Repository

```bash
git clone https://github.com/larrikin-coder/Excel-Mcp.git
cd Excel-Mcp
```

### 2. Set Up a Python Environment (Recommended)

It is recommended to use a virtual environment:

```bash
python -m venv venv
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install openpyxl
```

### 4. Run the Scripts

You can run the provided scripts directly. For example:

```bash
python create_sheet.py
python write_cell.py
```

To use the server or frontend (if implemented):

```bash
python mcp_server.py
python frontend.py
```

### 5. Using Your Own Excel Files

- Place your Excel files in the project directory or update the script paths as needed.
- Modify or extend the scripts to suit your workflow.

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

## Extending
Add new command modules in `modules.py` to support more Excel operations.

## Contributing
Contributions are welcome! Please open issues or pull requests on [GitHub](https://github.com/larrikin-coder/Excel-Mcp).

## License
MIT License
https://www.loom.com/share/982e32fa506f41e08b4ee6c794af2324