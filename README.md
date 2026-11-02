# PC Agent — AI-Powered Desktop Assistant

A local AI assistant for file management, browser automation, and Microsoft Office workflows.

PC Agent is a Python-based AI agent with a web interface built using FastAPI, HTML, CSS, and JavaScript. Powered by the Groq API, it interprets natural-language requests and uses registered tools to perform tasks involving files, web pages, and office documents.

The project combines LLM tool calling, browser automation, workspace file operations, and document generation in one application.

Repository: https://github.com/amit171ku/PC-AGENT

---

## Features

### AI Agent
- Interact with the agent through a browser-based interface.
- Execute supported actions through LLM tool calling.
- Configure Groq models in priority order.
- Automatically try eligible fallback models when a model is rate-limited, unavailable, or cannot format a tool call.
- Retry temporary network errors, timeouts, and Groq server errors (HTTP 5xx) before giving up.
- Start even when Groq is unreachable: model discovery happens at server startup, not at import time, and failures are reported instead of crashing the app.
- Reset conversations when starting a new task.

### File Management
- List, read, search, create, move, rename, and delete files using registered tools.
- Workspace path validation helps prevent access outside the workspace.
- Existing files are never silently overwritten: creating a file or moving onto an existing name is rejected.
- Deleted files go to a hidden `.trash` folder with a unique name, so earlier deleted files are not replaced.
- Uploaded files get a unique name prefix, so an upload cannot replace an existing file.
- Every action is written to `agent_log.txt` in the workspace.

### Browser Automation
- Search the web and open webpages.
- Extract webpage text, links, and tables.
- Interact with page elements using Playwright (click, fill, select, check, press keys).
- Assert elements and text on a page.
- Inspect browser console errors.
- Capture screenshots and export pages to PDF (saved inside the workspace, never overwriting existing files).
- Ask for approval before clicking, typing, selecting, checking, and pressing keys.
- Refuse to type into password, card, and other sensitive fields. The check looks at the real input type and attributes, not only the selector text.
- Block navigation to non-http(s) addresses and to local or private network addresses (localhost, `192.168.x.x`, `10.x.x.x`, cloud metadata addresses). Redirects are checked too.

### Microsoft Word Automation
- Create and read DOCX documents.
- Generate structured documents with headings and paragraphs.
- Apply supported formatting and design options.
- Create reports and other document-based outputs.

### Microsoft Excel Automation
- Create and edit XLSX workbooks.
- Work with worksheets and cell values.
- Use formulas and supported calculations.
- Apply formatting and create charts.
- Generate structured spreadsheet reports.

### Microsoft PowerPoint Automation
- Create PPTX presentations.
- Add and organize slide content.
- Apply supported layouts, formatting, and design options.
- Generate presentations for reports and project demonstrations.

> Available actions depend on the functions implemented and registered in the corresponding modules.

---

## Technology Stack

| Component | Technology |
|---|---|
| Programming language | Python |
| AI model API | Groq API |
| Backend framework | FastAPI |
| ASGI server | Uvicorn |
| Browser automation | Playwright |
| Browser engine | Chromium |
| Word documents | python-docx |
| Excel workbooks | openpyxl |
| PowerPoint presentations | python-pptx |
| PDF processing | pypdf |
| Charts and visuals | matplotlib |
| Frontend | HTML, CSS, JavaScript |
| Environment configuration | python-dotenv |
| Testing | pytest |

---

## Project Structure

```text
PC-AGENT/
├── main.py
├── common.py
├── requirements.txt
├── .gitignore
├── README.md
├── browser/
│   ├── __init__.py
│   └── browser_tools.py
├── office/
│   ├── __init__.py
│   ├── doc_tools.py
│   ├── word_tools.py
│   ├── excel_tools.py
│   └── ppt_tools.py
├── office_design/
│   ├── __init__.py
│   ├── design_tools.py
│   ├── word_design.py
│   ├── excel_design.py
│   └── ppt_design.py
├── static/
│   ├── index.html
│   ├── style.css
│   └── app.js
└── tests/
    ├── conftest.py
    ├── test_imports.py
    ├── test_common.py
    ├── test_models.py
    ├── test_files.py
    └── test_browser.py
```

> The repository on GitHub may still show the older flat layout (all modules in the root folder). `main.py` imports from the folders above, so keep the code and this layout in sync.

### Important files

| File | Responsibility |
|---|---|
| `main.py` | FastAPI application, agent loop, model setup and fallback, tool registration |
| `common.py` | Workspace paths, path validation, logging, and confirmation helpers |
| `browser/browser_tools.py` | Browser automation tools and browser safety checks |
| `office/word_tools.py` | Word document operations |
| `office/excel_tools.py` | Excel workbook operations |
| `office/ppt_tools.py` | PowerPoint operations |
| `office/doc_tools.py` | Additional document utilities |
| `office_design/` | Styling and design helpers for Word, Excel, and PowerPoint |
| `static/` | Web interface assets |
| `tests/` | Automated tests |
| `requirements.txt` | Python dependencies |

---

## Installation and Setup

### Prerequisites

Before starting, install:

- Python 3.10 or newer
- Git
- A Groq API key
- An internet connection for Groq API requests and web browsing

The application can run on Windows, macOS, and Linux.

### Step 1: Clone the repository

```bash
git clone https://github.com/amit171ku/PC-AGENT.git
cd PC-AGENT
```

### Step 2: Create a virtual environment

Windows PowerShell

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, you can use Command Prompt:

```bat
.venv\Scripts\activate.bat
```

macOS or Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### Step 3: Install dependencies

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Install Chromium for Playwright:

```bash
python -m playwright install chromium
```

### Step 4: Configure environment variables

Create a `.env` file in the same directory as `main.py`.

```dotenv
GROQ_API_KEY=your_groq_api_key
GROQ_MODELS=your_primary_model,your_fallback_model
GROQ_AUTO_FALLBACK=on
```

Replace the example values with valid model IDs available to your Groq account.

| Variable | Purpose |
|---|---|
| `GROQ_API_KEY` | Authenticates requests to the Groq API (required for chat). |
| `GROQ_MODELS` | Comma-separated list of preferred models in priority order. |
| `GROQ_MODEL` | Single-model alternative, used when `GROQ_MODELS` is not set. |
| `GROQ_AUTO_FALLBACK` | `on` (default) adds built-in fallback candidates. Set to `off` to use only your list. |
| `AGENT_WORKSPACE` | Folder the agent is allowed to work in. Default: `~/AgentWorkspace`. |
| `BROWSER_ALLOW_PRIVATE` | Optional. Set to `1` to let the browser open local/private network addresses (blocked by default). |
| `BROWSER_NO_SANDBOX` | Optional. Set to `1` only if Chromium cannot start otherwise (for example as root in Docker). The Chromium sandbox is enabled by default. |

Important: The code does not use `GROQ_MODEL_1`, `GROQ_MODEL_2`, or `GROQ_MODEL_3`. Use the variable names shown above.

Never commit your `.env` file or expose your API key publicly.

### Step 5: Run the application

```bash
python main.py
```

After startup, open:

http://127.0.0.1:8000

Check the terminal for the line `Models in priority order: [...]`. If Groq is unreachable or the API key is missing, the server still starts and prints a warning; chat requests will return a clear error until the problem is fixed.

Keep the terminal running while you use the application. To stop the server, press `Ctrl+C`.

---

## Example Use Cases

### File Management

- List the files in my workspace.
- Create a text file containing my daily tasks.
- Organize files into folders by extension.

### Browser Automation

- Search for recent Python best practices.
- Open a webpage and summarize its content.
- Extract links and tables from a webpage.
- Check a webpage for console errors.

### Microsoft Word

- Create a professional project report in Word.
- Generate a document with headings and a summary table.
- Apply consistent formatting to a supported DOCX document.

### Microsoft Excel

- Create a sales workbook with formulas and charts.
- Organize spreadsheet data into worksheets.
- Summarize key metrics from a workbook.

### Microsoft PowerPoint

- Create a presentation about an AI project.
- Organize content into structured slides.
- Apply consistent formatting to a presentation.

These are example requests. Successful execution depends on the tools registered in the current implementation, available input files, and the permissions granted to the agent.

---

## Groq Model Selection and Fallback

At server startup PC Agent builds a model priority list from `.env` and the models available to your Groq account, then checks which ones respond to a tool-calling request. If the model list cannot be fetched, it falls back to the models named in `.env`.

During a chat request, the agent switches to the next eligible model when:

- the current model is rate-limited (HTTP 429),
- the model does not support tool calling or is no longer available,
- the model fails to format a tool call.

Temporary problems (network errors, timeouts, HTTP 5xx) are retried up to three times. An invalid API key is reported immediately instead of being retried.

Switching models does not guarantee uninterrupted service or bypass Groq account limits.

For best results:

1. Use model IDs that are currently available to your account.
2. Confirm that the selected models support tool calling.
3. Check the startup logs for the model priority list and any skipped models.
4. Monitor API limits and usage through the Groq Console.

Groq resources: https://console.groq.com/docs/models

---

## Security Considerations

PC Agent can interact with files and websites, so security matters.

**Built-in safeguards**
- The server binds to `127.0.0.1` only.
- File paths are validated against the workspace folder.
- Creating, moving, and deleting files ask for your approval, and existing files are not overwritten.
- Browser actions that change a page (click, type, select, check, key press) ask for approval.
- Password, card, and similar fields are refused, based on the real input type and attributes.
- The browser refuses non-http(s) addresses and local/private network addresses.
- The browser's Chromium sandbox is on unless you turn it off with `BROWSER_NO_SANDBOX`.

**Your responsibilities**
- Keep API keys and credentials in the local `.env` file.
- Use a dedicated workspace for testing.
- Read every confirmation prompt before approving it.
- Treat webpage content and model-generated instructions as untrusted input.
- Review generated documents and browser actions before relying on them.
- Keep the server bound to localhost unless you have deliberately reviewed the implications of exposing it.

**Known limitations**
- The browser safety check covers page navigation. It does not inspect every image or script a page loads.
- The web interface does not yet use a session token. The server only accepts requests from its own localhost origin; treat it as a single-user local tool.

Path validation and confirmation prompts are useful safeguards, but they do not guarantee complete security. Review the code and permissions before using the agent with sensitive information.

---

## Testing

```bash
python -m pip install pytest
python -m pytest -v
```

The tests use a temporary workspace and do not need a Groq API key or internet access. They cover imports, workspace path validation, file tools, model fallback, and the browser URL safety check.

---

## Troubleshooting

| Problem | Possible solution |
|---|---|
| Missing Groq API key | Check that `.env` is beside `main.py` and contains `GROQ_API_KEY`. The server starts without it, but chat will not work. |
| "Invalid GROQ_API_KEY" | The key is wrong or revoked. Create a new key in the Groq Console. |
| "Cannot reach Groq" | Check your internet connection or firewall, then try again. |
| Model is unavailable | Verify the model ID and check availability in your Groq account. |
| Rate limit or HTTP 429 | Wait for the limit to reset, reduce request frequency, or add another eligible model to `GROQ_MODELS`. |
| Tool-calling error | Use a model that supports tool calling. |
| `ModuleNotFoundError` | Activate your virtual environment, reinstall `requirements.txt`, and check that imports match the folder layout. |
| Chromium fails to launch | Run `python -m playwright install chromium`. On Linux as root or in Docker, set `BROWSER_NO_SANDBOX=1`. |
| Browser says "Blocked: ... private network address" | This is the safety check. Set `BROWSER_ALLOW_PRIVATE=1` only if you really need local pages. |
| "File already exists" | File tools do not overwrite. Choose another file name. |
| Port 8000 is already in use | Stop the other process or change `PORT` in `main.py`. |
| Web interface does not load | Check the terminal logs and confirm that the `static/` directory exists. |
| Office file generation fails | Check input files, dependencies, permissions, and tool registration. |

---

## Development and Customization

### Add a new tool

1. Implement the Python function in the appropriate module.
2. Define its tool schema.
3. Register the function and schema in the agent.
4. Add a test for it in `tests/`.
5. Verify confirmation handling and file-path validation where applicable.

### Customize the interface

Modify the frontend files inside `static/`:

- `index.html` — page structure
- `style.css` — appearance and layout
- `app.js` — frontend behavior

### Modify AI behavior

Review the system instructions, model configuration, and tool-registration logic in `main.py`.

### Extend Office functionality

Add or modify functions in the relevant Word, Excel, PowerPoint, and design modules. Keep tool schemas synchronized with their Python implementations.

---

## Useful Resources

- GitHub Repository: https://github.com/amit171ku/PC-AGENT
- Groq Console: https://console.groq.com
- Groq Models: https://console.groq.com/docs/models
- FastAPI Documentation: https://fastapi.tiangolo.com/
- Playwright Documentation: https://playwright.dev/python/
- Python Documentation: https://docs.python.org/3/

---

## License

No license file is currently listed in the repository. Until a license is added, do not assume the project is available for unrestricted reuse, modification, or redistribution.