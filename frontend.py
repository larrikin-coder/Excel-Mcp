import streamlit as st
import requests
import os
import base64
from openpyxl import Workbook
from dotenv import load_dotenv

load_dotenv()

MCP_SERVER_URL = os.getenv("MCP_SERVER_URL")

st.set_page_config(page_title="Exl LLM 🚀", page_icon="📊")
st.title("Exl LLM 🚀")

if 'history' not in st.session_state:
    st.session_state.history = []

if 'uploaded_filename' not in st.session_state:
    st.session_state.uploaded_filename = "uploaded_file.xlsx"

st.subheader("✨ Create New Excel File")

if st.button("📄 Create New Blank Excel"):
    wb = Workbook()
    wb.save(st.session_state.uploaded_filename)
    st.success("New blank Excel created!")

if os.path.exists(st.session_state.uploaded_filename):
    st.subheader("Ask something about Excel ✍️")
    user_prompt = st.text_input("Enter your prompt (e.g., Create a sheet called Finance)")

    if st.button("💬 Send Prompt"):
        if user_prompt:
            payload = {"prompt": user_prompt}
            response = requests.post(f"{MCP_SERVER_URL}/chat", json=payload)

            if response.status_code == 200:
                result = response.json()
                st.session_state.history.append((user_prompt, result))

                # ✅ Decode and save file locally
                if "results" in result and len(result["results"]) > 0:
                    last = result["results"][-1]
                    if "file" in last:
                        file_data = base64.b64decode(last["file"])
                        with open(st.session_state.uploaded_filename, "wb") as f:
                            f.write(file_data)

                st.success("✅ Prompt processed!")
            else:
                st.error(f"❌ Error: {response.status_code} {response.text}")

    # Chat history
    st.subheader("Chat History 💬")
    for prompt, result in st.session_state.history[::-1]:
        st.markdown(f"**You:** {prompt}")
        st.json(result)

    # Download button
    if os.path.exists(st.session_state.uploaded_filename):
        with open(st.session_state.uploaded_filename, "rb") as f:
            st.download_button(
                label="📥 Download Updated Excel",
                data=f,
                file_name="updated_excel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
