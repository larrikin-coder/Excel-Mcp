import os
import json
import re
import base64
from openpyxl import load_workbook, Workbook


def file_to_base64(filepath: str) -> str:
    """Convert a file to base64 string."""
    with open(filepath, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")
