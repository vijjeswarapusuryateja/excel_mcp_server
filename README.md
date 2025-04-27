# Excel MCP Server 🚀

This project implements a lightweight **Model Context Protocol (MCP)** server for Excel automation, powered by **OpenAI's ChatGPT (GPT-4o)**.

You can create, edit, and automate Excel files through natural language prompts, with GPT translating your intent into structured tool calls.

---

## ✨ Features

- Lightweight FastAPI backend exposing Excel operations as MCP tools
- Natural language prompt handling with GPT-4o orchestration
- Multi-tool call execution (supports workflows like creating sheets + writing cells)
- Streamlit frontend to interact easily
- Minimal setup — no complex SDKs required

---

## 📦 Project Structure

| File | Purpose |
|:--|:--|
| `excel_mcp_server.py` | FastAPI backend with Excel MCP tools and GPT integration |
| `frontend.py` | Streamlit frontend for chatting with Excel |
| `requirements.txt` | (Coming soon) List of Python dependencies |

---

## 🚀 How to Run Locally

1. Clone this repo:

```bash
git clone https://github.com/vijjeswarapusuryateja/excel_mcp_server.git
cd excel_mcp_server
```

2. Create a virtual environment:

```bash
python3 -m venv venv
source venv/bin/activate  # Mac/Linux
venv\Scripts\activate  # Windows
```

3. Install required packages:

```bash
pip install -r requirements.txt
```

4. Run the backend server:

```bash
python excel_mcp_server.py
```

5. In a new terminal, run the frontend:

```bash
streamlit run frontend.py
```

---

## 📚 Related Article

I wrote a full Medium article explaining the architecture, how I built this project, and what I learned about MCP:

👉 [Read the article here](https://medium.com/@surya.vijjeswarapu/how-i-built-a-lightweight-excel-mcp-server-using-openais-chatgpt-and-understood-model-context-544a539d0f07)

---

## 🧐 Future Improvements

- Dockerize backend + frontend
- Add batch writing tools
- Add style formatting (bold, colors)
- Support uploading/downloading Excel files through API

---

## 📜 License

MIT License

---

## 🌟 If you found this project useful, give it a ⭐ star and feel free to fork/extend!
