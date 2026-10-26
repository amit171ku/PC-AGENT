# PC Agent — AI-Powered Desktop Assistant

PC Agent is a local AI agent with a modern web interface, powered by the Groq API. It can manage files in a sandboxed workspace, browse the web using Chromium, and automate document workflows for Word, Excel, and PowerPoint.

## Features

### AI Agent
- Natural-language commands through a browser-based interface.
- Tool calling for file management, web browsing, and office document operations.
- Automatic selection and fallback across three configurable Groq models.
- Conversation reset to help reduce token usage.

### File Management
- List, read, search, create, rename, move, and delete files within the configured workspace.
- Safe deletes using a hidden `.trash` directory.
- Action logging through `agent_log.txt`.
- Workspace path validation to help prevent access outside the permitted directory.

### Web Browsing
- Search the web and open webpages.
- Read page content and interact with supported website elements using Playwright and Chromium.
- Approval prompts for supported actions such as clicking and typing.
- Visible Chromium window for browser automation.

### Microsoft Office Automation

**Word — DOCX**
- Create and read Word documents.
- Edit document content and apply formatting.
- Generate structured reports with headings, paragraphs, and tables.

**Excel — XLSX**
- Create and edit Excel workbooks.
- Write formulas and organize worksheets.
- Apply cell formatting, tables, and charts.
- Build summaries and reports from structured data.

**PowerPoint — PPTX**
- Generate presentations with structured slide layouts.
- Add titles, text, images, shapes, and diagrams.
- Create charts and visual summaries when relevant data is available.
- Apply consistent typography, spacing, colors, and slide styling.

Office capabilities depend on the tools and functions implemented in the corresponding modules.

### Three-Model Groq Fallback

Configure three Groq models to improve resilience when the primary model cannot complete a request.

Expected behavior:
1. Try the primary model.
2. Switch to the second model when a retryable error occurs.
3. Try the third model if the earlier attempts fail.
4. Report a useful error if all configured models fail.

Fallback should handle supported transient errors, rate limits, and model/tool compatibility issues. Invalid credentials and malformed requests should be reported instead of retried indefinitely.

Model availability and tool-calling support depend on Groq's current model lineup and your account limits. Switching models does not guarantee that rate limits will be avoided.

## Technology Stack

| Component | Technologies |
|---|---|
| Backend | Python, FastAPI, Uvicorn |
| AI and tool calling | Groq API |
| Browser automation | Playwright, Chromium |
| Word documents | `python-docx` |
| Excel workbooks | `openpyxl` |
| PowerPoint presentations | `python-pptx` |
| Frontend | HTML, CSS, JavaScript |
| Configuration | `python-dotenv` |

## Requirements

- Python 3.10 or newer
- Groq API key
- Windows, macOS, or Linux
- Internet access for Groq API calls and web browsing

## Project Structure

```text
PC-AGENT/
├── main.py
├── common.py
├── word_tools.py
├── excel_tools.py
├── ppt_tools.py
├── doc_tools.py
├── requirements.txt
├── .env.example
├── .gitignore
├── README.md
└── static/
    ├── index.html
    ├── style.css
    └── app.js
```

The structure above is illustrative. Your actual repository may contain different modules.

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/amit171ku/PC-AGENT.git
cd PC-AGENT
```

### 2. Create a virtual environment

**Windows PowerShell**

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
```

**macOS or Linux**

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies

```bash
python -m pip install -r requirements.txt
playwright install chromium
```

If your Office modules require these libraries, ensure `requirements.txt` includes:

```text
groq
playwright
python-dotenv
fastapi
uvicorn[standard]
python-docx
openpyxl
python-pptx
```

### 4. Configure environment variables

Create a local `.env` file in the same directory as `main.py`.

```dotenv
GROQ_API_KEY=your_groq_api_key
GROQ_MODEL_1=your_primary_model
GROQ_MODEL_2=your_fallback_model
GROQ_MODEL_3=your_final_fallback_model
```

Replace the model placeholders with currently available Groq models that support the tool-calling functionality your agent requires.

The application must explicitly read these variables and implement the fallback logic. Merely adding them to `.env` does not enable automatic switching.

**Security:** Never commit your real `.env` file or publish API keys.

### 5. Run the agent

```bash
python main.py
```

Open the local URL printed in the terminal, typically:

http://127.0.0.1:8000

Stop the server with `Ctrl+C`.

## Example Commands

### File Management
- `List all files in my workspace.`
- `Create a todo.txt file with three tasks.`
- `Organize files into folders by extension.`

### Web Browsing
- `Search for recent Python best practices and save a summary.`
- `Open Wikipedia and summarize the page.`

### Word
- `Create a professional project report in Word.`
- `Format this DOCX file with headings and tables.`

### Excel
- `Create an Excel sales report with formulas and charts.`
- `Analyze this workbook and summarize the key metrics.`

### PowerPoint
- `Create a professional 7-slide AI project presentation.`
- `Generate an architecture diagram and workflow slide.`
- `Create a presentation using a consistent theme and layout.`

### Model Fallback
- `Try the configured Groq models automatically if the primary model hits a rate limit.`

The agent can execute these tasks only to the extent supported by its implemented tools and permissions.

## Security

- Bind the local server to `127.0.0.1` by default.
- Validate file paths against the configured workspace.
- Keep API keys and private configuration out of version control.
- Require backend-enforced approval for sensitive actions.
- Validate browser interactions and reject disallowed sensitive form inputs.
- Use a temporary workspace when testing file operations.
- Treat web content and model-generated instructions as untrusted input.
- Review generated documents and browser actions before using them in important workflows.

Approval prompts are useful safeguards, but they do not replace backend validation and access controls.

## Troubleshooting

| Problem | Solution |
|---|---|
| Missing API key | Verify `.env` location and variable names. |
| Rate limit or HTTP 429 | Apply backoff and switch to an eligible fallback model. |
| Model does not support tools | Configure a model that supports the required tool-calling interface. |
| `ModuleNotFoundError` | Activate the virtual environment and install requirements. |
| Chromium fails to launch | Run `playwright install chromium`. |
| Office file generation fails | Check the relevant library, input data, file permissions, and module registration. |
| Static assets return 404 | Verify the FastAPI static-file mount and directory path. |
| Address already in use | Stop the other server process or configure a different port. |

## Customization

- **Add tools:** Implement the tool function and register it with the agent.
- **Change AI behavior:** Update the system prompt and tool instructions.
- **Customize Word, Excel, and PowerPoint:** Modify the relevant office modules.
- **Change the interface:** Edit files in `static/`.
- **Configure model fallback:** Update `.env` and verify the fallback implementation in the backend.

## Resources

- GitHub Repository: https://github.com/amit171ku/PC-AGENT
- Groq Console: https://console.groq.com
- Groq Models: https://console.groq.com/docs/models
- FastAPI: https://fastapi.tiangolo.com/
- Playwright: https://playwright.dev/python/

## License

No license has been specified. Add a license file if you want to define how others may use, modify, and distribute the project.
