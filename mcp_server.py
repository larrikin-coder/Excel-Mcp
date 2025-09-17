import os
import json
import traceback
from dotenv import load_dotenv

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware

import google.generativeai as genai
from openai import AsyncOpenAI

from create_sheet import create_sheet
from write_cell import write_cell

import uvicorn

load_dotenv()

api_key = os.getenv("OPENAI_API_KEY")
openai_client = AsyncOpenAI(api_key=api_key)

genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
model = genai.GenerativeModel('gemini-2.5-flash')

# -------------------------------
# FastAPI App + CORS
# -------------------------------
app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # restrict later
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -------------------------------
# MCP Infrastructure
# -------------------------------
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

    async def call(self, client, res):
        try:
            tool_call = res['tool_calls'][0]
            tool_name = tool_call['function']['name']
            tool_args_raw = tool_call['function']['arguments']

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


mcp_handler = Middleware(
    tool_functions=[
        ToolFunction.from_function(write_cell),
        ToolFunction.from_function(create_sheet),
    ]
)

# -------------------------------
# Endpoints
# -------------------------------
@app.post("/mcp")
async def handle_mcp(request: Request):
    try:
        body = await request.json()
        result = await mcp_handler.call(openai_client, body)
        return JSONResponse(content=result)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

def extract_json_from_markdown(raw_text: str) -> str:
    raw_text = raw_text.strip()
    if raw_text.startswith("```"):
        lines = raw_text.splitlines()
        if len(lines) >= 3:
            return "\n".join(lines[1:-1])
    return raw_text

@app.post("/chat")
async def chatgpt(request: Request):
    try:
        data = await request.json()
        prompt = data.get("prompt")
        default_filepath = "uploaded_file.xlsx"

        GEMINI_SYSTEM_PROMPT = """
        You are an Excel control agent. You are only allowed to respond with a JSON structure that represents a tool/function call.

        ❗ IMPORTANT: Always include all required arguments.
        - For create_sheet: filepath, sheet_name
        - For write_cell: filepath, sheet_name, cell, value

        Format:
        {
        "function": {
            "name": "tool_name",
            "arguments": {
            "param1": "value1",
            "param2": "value2",
            "filepath": "uploaded_file.xlsx"
            }
        }
        }
        """

        formatted_prompt = f"{GEMINI_SYSTEM_PROMPT.strip()}\n\nUser input: {prompt}"

        response = model.generate_content(
            contents=[{"role": "user", "parts": [formatted_prompt]}]
        )

        reply = extract_json_from_markdown(response.text.strip())

        try:
            tool_data = json.loads(reply)
            if not isinstance(tool_data, list):
                tool_data = [tool_data]
        except Exception as e:
            return JSONResponse(
                status_code=500,
                content={"error": f"Failed to parse Gemini response as JSON: {e}", "raw_response": reply}
            )

        all_responses = []
        for tool in tool_data:
            if "function" in tool:
                args = tool["function"]["arguments"]

                # Always ensure filepath
                args.setdefault("filepath", default_filepath)

                # Add safe defaults
                if tool["function"]["name"] == "write_cell":
                    args.setdefault("sheet_name", "Sheet1")
                    args.setdefault("cell", "A1")
                    args.setdefault("value", "")
                elif tool["function"]["name"] == "create_sheet":
                    args.setdefault("sheet_name", "Sheet1")

                mcp_response = await mcp_handler.call(None, {"tool_calls": [tool]})
                all_responses.append(mcp_response)

        return {"results": all_responses}

    except Exception as e:
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"error": str(e)})

# -------------------------------
# Run (Render dynamic port)
# -------------------------------
if __name__ == "__main__":
    port = int(os.getenv("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)
