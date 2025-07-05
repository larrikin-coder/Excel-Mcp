import os
import sys
import json
import re

from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv("OPENAI_API_KEY")

import uvicorn
from fastapi import FastAPI, Request
from openai import AsyncOpenAI
from openpyxl import Workbook,load_workbook

# wb =  Workbook()





class MCPServer:
    def __init__(self,name,description,parameters,function):
        self.name = name
        self.description = description
        self.parameters = parameters
        self.function = function
        
        
    def run(self,**kwargs):
        return self.function(**kwargs)
    
    
    @staticmethod
    def from_function(fn):
        import inspect
        
        sig = inspect.signature(fn)
        params = {}
        for name,param in sig.parameters.items():
            params[name] = {"type":"string","description": ""}
        return MCPServer(
            name=fn.__name__,
            description=fn.__doc__ or "",
            parameters=params,
            function=fn
            )

    
    def openai_tools(self):
        tools = [{
            "type": "function",
            "name": self.name,
            "description": self.description,
            "parameters": {
                "type": "object",
                "properties": self.parameters,
                "required": list(self.parameters.keys())
            }
        }]
        return tools

class ToolFunction:
    @staticmethod
    def from_function(fn):
        return MCPServer.from_function(fn)
    
    
    
class Middleware:
    def __init__(self,tool_functions):
        self.tool_functions = tool_functions

    async def call(self,res):
        tool_call = res['tool_calls'][0]
        tool_name = tool_call['function']['name']
        tool_args_raw = tool_call['function']['arguments']
         
        if not isinstance(tool_args_raw,str):
            raise ValueError
        
        if isinstance(tool_args_raw,str):
            try:
                res = json.loads(tool_args_raw)
            except Exception:
                res =  tool_args_raw
        
         
        for tool in self. tool_functions:
            if tool.name == tool_name:
                return tool.run(**res)
         
        raise Exception(f"No tool found for {tool_name}")
        
        
    def openai_tools(self):
        return [tool.openai_tool() for tool in self.tool_functions]


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


def write_cell(filepath: str, sheet_name: str, cell: str, value: str):
    """Write a value to a specific cell. Create sheet if missing. Handle spaces, cases safely."""
    wb = load_workbook(filepath)
    
    # Match ignoring cases and spaces
    clean_sheet_names = {s.strip().lower(): s for s in wb.sheetnames}
    requested_sheet_clean = sheet_name.strip().lower()
    
    if requested_sheet_clean in clean_sheet_names:
        real_sheet_name = clean_sheet_names[requested_sheet_clean]
        ws = wb[real_sheet_name]
    else:
        # Create sheet if not found
        ws = wb.create_sheet(title=sheet_name)
    
    ws[cell] = value
    wb.save(filepath)
    return {"message": f"Value '{value}' written to {sheet_name}:{cell}"}


# Initialize FastAPI app
app = FastAPI()

# Register MCP Tools
mcp_handler = Middleware(
    tool_functions=[
        ToolFunction.from_function(write_cell),
        ToolFunction.from_function(create_sheet),
    ]
)

# Route 1: Direct MCP Call
@app.post("/mcp")
async def handle_mcp(request: Request):
    body = await request.json()
    response = await mcp_handler.call(body)
    return response

# Main server runner
if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
             
         
        