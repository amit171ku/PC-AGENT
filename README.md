# PC Agent

A local AI agent with a web interface. Type a command in your browser and the agent can manage files in a sandboxed workspace folder and search or browse the web, powered by the Groq API.

## Features

- **File management:** list, read, search, create, rename, move and delete files inside one workspace folder
- **Web browsing:** search the web, open pages, read page text, click and type (a real Chromium window opens so you can watch)
- **Approval prompts:** writing, moving, deleting, clicking and typing all ask Allow/Deny first
- **Safe deletes:** deleted files go to a hidden `.trash` folder instead of being erased
- **Action log:** every tool call is saved to `agent_log.txt`
- **Auto model pick:** tests your Groq models at startup and uses one that supports tool calling
- **Modern UI:** sidebar with quick actions and live workspace file list, light/dark theme, mobile friendly

## Requirements

- Python 3.10 or newer
- A free Groq API key from https://console.groq.com
- Windows, macOS or Linux

## Project structure

```
pc-agent/
├── main.py              FastAPI backend + agent logic
├── requirements.txt
├── .env                 your API key (never share or commit this)
├── .gitignore
├── README.md
├── venv/                virtual environment (created by you)
└── static/
    ├── index.html       page layout
    ├── style.css        styling
    └── app.js           frontend logic
```

## Setup

Run these from the project folder (PowerShell on Windows).

**1. Create and activate a virtual environment**

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
```

macOS / Linux:

```bash
python3 -m venv venv
source venv/bin/activate
```

If PowerShell blocks the activation script, run this once and try again:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

**2. Install packages**

`requirements.txt`:

```
groq
playwright
python-dotenv
fastapi
uvicorn[standard]
```

```powershell
pip install -r requirements.txt
playwright install chromium
```

**3. Create the `.env` file**

```
GROQ_API_KEY=your_key_here
GROQ_MODEL=openai/gpt-oss-120b
```

- No spaces around `=` and no quotes
- `GROQ_MODEL` is optional. If it is missing or does not support tools, the app picks another model automatically.

**4. Create `.gitignore`**

```
.env
venv/
__pycache__/
```

## Run

```powershell
cd "path\to\pc-agent"
venv\Scripts\Activate.ps1
python main.py
```

The terminal prints `Using model: ...` and `Uvicorn running on http://127.0.0.1:8000`. Open **http://127.0.0.1:8000** in your browser.

Stop the server with `Ctrl+C`.

## Usage

Type a command and press Enter (Shift+Enter for a new line). Examples:

- `List all files in my workspace`
- `Create a file called todo.txt with 3 tasks for today`
- `Sort the files in my workspace into folders by file type`
- `Search the web for the best Python tips and save a short summary in tips.txt`
- `Open wikipedia.org and summarize the homepage`

The sidebar has quick-action buttons for common tasks and a live list of workspace files. Use **New chat** to clear the conversation, which also saves tokens.

## Where files go

By default the agent works only inside:

```
C:\Users\<you>\AgentWorkspace
```

It is created automatically on first run and also holds `agent_log.txt` and the hidden `.trash` folder. To use a different folder, edit this line in `main.py`:

```python
ROOT = Path.home() / "AgentWorkspace"
# or keep it inside the project:
ROOT = BASE / "workspace"
```

Start with a test folder before pointing the agent at real documents.

## Safety

- The agent can only touch files inside `ROOT`. Paths that escape it are rejected.
- The server binds to `127.0.0.1`, so only your PC can reach it. Do not change it to `0.0.0.0`.
- The WebSocket rejects connections from other websites.
- The agent refuses to type into fields named like `password`, `card` or `cvv`.
- Unanswered approval prompts are denied after 2 minutes.
- Never share your `.env` file. If a key leaks, delete it in the Groq console and create a new one.

## Troubleshooting

| Problem | Fix |
|---|---|
| `KeyError: 'GROQ_API_KEY'` | `.env` is missing, in the wrong folder, or misspelled. It must sit next to `main.py`. |
| `tool calling is not supported with this model` | The model does not support tools. Pick a chat model from your Groq model list (not names containing `guard`, `whisper` or `tts`). |
| `ModuleNotFoundError` | Activate the venv, then run `pip install -r requirements.txt`. |
| Browser window never opens | Run `playwright install chromium`. |
| `page closed` errors | The Chromium window was closed. Update `page()` in `main.py` to reopen it automatically, then restart the server. |
| Page has no styling (`/static/style.css` shows Not Found) | Add `app.mount("/static", StaticFiles(directory=BASE / "static"), name="static")` after `app = FastAPI(...)` and restart the server. |
| Workspace panel says "File list unavailable" | The `/files` route is missing from `main.py`, or the server was not restarted. |
| `Address already in use` | Another copy is running. Stop it, or change `PORT` in `main.py`. |
| `rate_limit` or 429 errors | Groq free limits are small. Wait a minute or click **New chat**. |
| Old UI still shows after editing | Hard refresh with `Ctrl+F5`. Changes to `main.py` need a server restart. |

## Customizing

- **Add a tool:** write a function in `main.py`, add it to `FUNCS` and add a `tool(...)` entry to `TOOLS`.
- **Change behavior:** edit the `SYSTEM` prompt in `main.py`.
- **Change the look:** edit `static/style.css`. Colors are CSS variables at the top of the file.
- **Reduce token use:** lower the `[:3000]` text limits in the read and browser tools, or trim the history length in `worker()`.

## Notes

- Groq's model lineup changes often. Check https://console.groq.com/docs/models for current tool-use models.
- Free-tier limits are low, so keep chats short.