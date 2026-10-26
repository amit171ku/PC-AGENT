# PC Agent

A local AI agent with a modern web interface, powered by the Groq API. Give it commands through your browser to manage files inside a sandboxed workspace and search or browse the web.

## Features

- **File Management:** List, read, search, create, rename, move, and delete files within a configured workspace.
- **Web Browsing:** Search the web, open pages, read page content, and interact with websites using Playwright and Chromium.
- **Approval Prompts:** Request Allow/Deny approval before supported file and browser actions.
- **Safe Deletes:** Move deleted files to a hidden `.trash` folder instead of permanently deleting them.
- **Action Logging:** Record tool calls in `agent_log.txt`.
- **Automatic Model Selection:** Test configured Groq models and select a compatible model for tool calling.
- **Modern UI:** Responsive interface with a sidebar, quick actions, workspace file list, and light/dark themes.

## Requirements

- Python 3.10 or newer
- A Groq API key: https://console.groq.com
- Windows, macOS, or Linux

## Project Structure

```text
pc-agent/
├── main.py
├── requirements.txt
├── .env.example
├── .gitignore
├── README.md
└── static/
    ├── index.html
    ├── style.css
    └── app.js
```

The virtual environment, actual `.env` file, local workspace, logs, and trash folder are created or maintained locally and should not be committed.

## Installation

### 1. Clone the repository

```powershell
git clone https://github.com/amit171ku/PC-AGENT.git
cd PC-AGENT
```

### 2. Create a virtual environment

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
```

On macOS or Linux:

```bash
python3 -m venv venv
source venv/bin/activate
```

If PowerShell blocks virtual environment activation, use:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### 3. Install dependencies

```bash
python -m pip install -r requirements.txt
playwright install chromium
```

### 4. Configure environment variables

Create a local `.env` file next to `main.py`:

```dotenv
GROQ_API_KEY=your_groq_api_key
GROQ_MODEL=openai/gpt-oss-120b
```

The model setting is optional if the application supports automatic model selection. Confirm that the selected model is currently available to your Groq account and supports tool calling.

**Never commit your real `.env` file or expose your API key.**

### 5. Run the application

```bash
python main.py
```

Open the local URL printed in the terminal, typically:

http://127.0.0.1:8000

Stop the server with `Ctrl+C`.

## Usage Examples

Try commands such as:

- `List all files in my workspace`
- `Create a file called todo.txt with three tasks for today`
- `Organize my workspace files into folders by file type`
- `Search the web for Python best practices and save a summary to tips.txt`
- `Open Wikipedia and summarize its homepage`

Use **New Chat** to start a fresh conversation and reduce the amount of chat history sent to the model.

## Workspace and File Safety

By default, the agent uses:

```text
C:\Users\<your-username>\AgentWorkspace
```

The workspace is created automatically. It can contain generated files, the action log, and the hidden trash folder.

To change the workspace, configure the `ROOT` path in `main.py`, or use an environment variable if the application supports it.

**Security considerations:**

- Restrict file operations to the configured workspace and validate resolved paths.
- Keep the web server bound to `127.0.0.1` unless a secure deployment architecture is implemented.
- Validate WebSocket origins and enforce action-specific approval checks on the backend.
- Deny sensitive browser actions and time out unanswered approval requests.
- Test file operations in a temporary workspace before using real documents.
- Do not treat approval prompts alone as a complete security boundary.

## Troubleshooting

| Problem | Suggested fix |
|---|---|
| `KeyError: GROQ_API_KEY` | Check that `.env` exists beside `main.py` and contains the correct variable name. |
| Model does not support tool calling | Select a currently available Groq model that supports tool use. |
| `ModuleNotFoundError` | Activate the virtual environment and reinstall requirements. |
| Chromium does not open | Run `playwright install chromium`. |
| Page has no styling | Check the FastAPI static-file mount and browser console. |
| Workspace files are unavailable | Verify the file-list endpoint and restart the server. |
| Address already in use | Stop the other server process or configure a different port. |
| Rate limit or HTTP 429 | Wait for the limit to reset, reduce requests, or use an eligible alternative model. |
| Old UI still appears | Hard refresh with `Ctrl+F5`. Restart the backend after Python changes. |

## Customization

- **Add tools:** Implement a tool function and register it in the application's tool definitions.
- **Change agent behavior:** Update the system prompt in `main.py`.
- **Customize the interface:** Edit `static/style.css`.
- **Reduce token usage:** Limit browser output, file contents, and conversation history.

## Resources

- Groq Console: https://console.groq.com
- Groq Models Documentation: https://console.groq.com/docs/models
- Playwright Documentation: https://playwright.dev/python/
- FastAPI Documentation: https://fastapi.tiangolo.com/

## License

No license has been specified yet. Add a `LICENSE` file before granting others explicit permission to reuse, modify, or redistribute the project.
