import streamlit as st
import requests
import os
from openpyxl import Workbook

# MCP server URL
MCP_SERVER_URL = "http://127.0.0.1:8000"

st.set_page_config(page_title="Excel Copilot 🚀", page_icon="📊")
st.title("Excel Copilot 🚀")

# Session state for history and file
if 'history' not in st.session_state:
    st.session_state.history = []

if 'uploaded_filename' not in st.session_state:
    st.session_state.uploaded_filename = None

st.subheader("✨ Create New Excel File")

if st.button("📄 Create New Blank Excel"):
    temp_filename = "uploaded_file.xlsx"
    wb = Workbook()
    wb.save(temp_filename)
    st.session_state.uploaded_filename = temp_filename
    st.success("✅ New blank Excel created!")

# Show prompt section if file is ready
if st.session_state.uploaded_filename:

    st.subheader("Ask something about Excel ✍️")
    user_prompt = st.text_input("Enter your prompt (e.g., Create a sheet called Finance)")

    if st.button("💬 Send Prompt"):
        if user_prompt:
            payload = {"prompt": user_prompt}
            response = requests.post(f"{MCP_SERVER_URL}/ask", json=payload)

            if response.status_code == 200:
                result = response.json()
                st.session_state.history.append((user_prompt, result))
                st.success("✅ Prompt processed!")
            else:
                st.error(f"❌ Error: {response.status_code}")

    # Chat history display
    st.subheader("Chat History 💬")
    for prompt, result in st.session_state.history[::-1]:
        st.markdown(f"**You:** {prompt}")
        st.json(result)

    # Download updated file
    if os.path.exists(st.session_state.uploaded_filename):
        with open(st.session_state.uploaded_filename, "rb") as f:
            st.download_button(
                label="📥 Download Updated Excel",
                data=f,
                file_name="updated_excel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
