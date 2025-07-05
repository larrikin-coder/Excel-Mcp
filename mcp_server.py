import os
import sys
import json
import re

from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv("OPENAI_API_KEY")