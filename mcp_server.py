import os
import json
from dotenv import load_dotenv

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from openai import AsyncOpenAI
from openpyxl import Workbook, load_workbook

import uvicorn


from write_cell import *
from create_sheet import *
load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")


class MCPServer:
    def __init__(self, name, description, parameters, function):
        self.name = name
        self.description = description
        self.parameters = parameters
        self.function = function

    def run(self, **kwargs):
        return self.function(**kwargs)

    @staticmethod
    def from_function(fn):
        import inspect
        sig = inspect.signature(fn)
        params = {
            name: {"type": "string", "description": ""}
            for name in sig.parameters
        }
        return MCPServer(
            name=fn.__name__,
            description=fn.__doc__ or "",
            parameters=params,
            function=fn
        )

    def openai_tool(self):
        return {
            "type": "function",
            "function": {
                "name": self.name,
                "description": self.description,
                "parameters": {
                    "type": "object",
                    "properties": self.parameters,
                    "required": list(self.parameters.keys())
                }
            }
        }


class ToolFunction:
    @staticmethod
    def from_function(fn):
        return MCPServer.from_function(fn)


class Middleware:
    def __init__(self, tool_functions):
        self.tool_functions = tool_functions

    async def call(self, res):
        try:
            tool_call = res['tool_calls'][0]
            tool_name = tool_call['function']['name']
            tool_args_raw = tool_call['function']['arguments']

            # Handle both string or dict input
            if isinstance(tool_args_raw, str):
                tool_args = json.loads(tool_args_raw)
            elif isinstance(tool_args_raw, dict):
                tool_args = tool_args_raw
            else:
                raise ValueError("Invalid arguments format")

            for tool in self.tool_functions:
                if tool.name == tool_name:
                    return tool.run(**tool_args)

            raise Exception(f"No tool found for {tool_name}")

        except Exception as e:
            print(f"Middleware Error: {e}")
            raise e

    def openai_tools(self):
        return [tool.openai_tool() for tool in self.tool_functions]


# def create_sheet(filepath: str, sheet_name: str):
#     """Create a new sheet, create file if it doesn't exist."""
#     if not os.path.exists(filepath):
#         wb = Workbook()
#         wb.create_sheet(title=sheet_name)
#         default_sheet = wb["Sheet"]
#         wb.remove(default_sheet)
#     else:
#         wb = load_workbook(filepath)
#         wb.create_sheet(title=sheet_name)

#     wb.save(filepath)
#     return {"message": f"Sheet '{sheet_name}' created in '{filepath}'."}

# def write_cell(filepath: str, sheet_name: str, cell: str, value: str):
#     """Write a value to a specific cell. Create sheet if missing. Handle spaces, cases safely."""
#     wb = load_workbook(filepath)

#     clean_sheet_names = {s.strip().lower(): s for s in wb.sheetnames}
#     requested_clean = sheet_name.strip().lower()

#     if requested_clean in clean_sheet_names:
#         ws = wb[clean_sheet_names[requested_clean]]
#     else:
#         ws = wb.create_sheet(title=sheet_name)

#     ws[cell] = value
#     wb.save(filepath)
#     return {"message": f"Value '{value}' written to {sheet_name}:{cell}"}

app = FastAPI()

mcp_handler = Middleware(
    tool_functions=[
        ToolFunction.from_function(write_cell),
        ToolFunction.from_function(create_sheet),
    ]
)

@app.post("/mcp")
async def handle_mcp(request: Request):
    try:
        body = await request.json()
        result = await mcp_handler.call(body)
        return JSONResponse(content=result)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
