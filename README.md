# PC Agent — Local AI Desktop Assistant

A local AI-powered desktop assistant that can manage files, automate browser tasks, and work with Microsoft Office documents using natural language.

PC Agent uses **Python, FastAPI, Groq LLMs, Playwright, and modular tool-calling** to turn user instructions into controlled actions.

## Features

### AI Agent
- Uses Groq models to understand natural-language instructions.
- Executes tasks through function/tool calling.
- Supports model selection and automatic fallback.
- Maintains a multi-step agent execution loop.

### File Management
- List and search files in the agent workspace.
- Read supported file formats.
- Create, move, and delete files.
- Move deleted files into a `.trash` directory.
- Validate file paths to prevent access outside the workspace.

### Browser Automation
- Search the web and open pages.
- Extract page text, links, and tables.
- Click elements and interact with forms.
- Fill fields, select options, and interact with checkboxes.
- Capture screenshots and export pages to PDF.
- Request approval before sensitive browser actions.

### Microsoft Office Automation
- Create and edit Word documents.
- Work with Excel spreadsheets.
- Generate PowerPoint presentations.
- Read supported document formats.
- Apply document designs through separate design modules.

### Safety Controls
- Workspace path validation.
- Approval prompts for controlled actions.
- Safer deletion through a trash directory.
- Browser URL checks.
- Checks for sensitive form fields.
- Action confirmation before selected operations.

## Tech Stack

| Component | Technology |
|---|---|
| Backend | Python, FastAPI |
| AI Model | Groq API |
| Frontend | HTML, CSS, JavaScript |
| Browser Automation | Playwright |
| Word | python-docx |
| Excel | openpyxl |
| PowerPoint | python-pptx |
| PDF | pypdf |
| Testing | pytest |
| Server | Uvicorn |

## Project Architecture

```text
PC-AGENT/
│
├── browser/
│   └── browser_tools.py
│
├── office/
│   ├── doc_tools.py
│   ├── word_tools.py
│   ├── excel_tools.py
│   └── ppt_tools.py
│
├── office_design/
│   ├── design_tools.py
│   ├── word_design.py
│   ├── excel_design.py
│   └── ppt_design.py
│
├── static/
│   ├── app.js
│   ├── index.html
│   ├── style.css
│   ├── markdown.js
│   └── markdown.css
│
├── common.py
├── main.py
├── test_basic.py
├── requirements.txt
├── README.md
└── .github/
    └── workflows/
        └── main.yml
```

### How It Works

1. The user sends an instruction through the web interface.
2. FastAPI receives the request.
3. The Groq model interprets the instruction.
4. The model selects the appropriate tool.
5. The backend executes the tool through the registered Python functions.
6. If an action requires approval, the agent waits for confirmation.
7. The result is returned to the frontend.
8. The agent continues until the task is complete or the execution limit is reached.

## Installation

### Requirements

- Python 3.10 or newer
- Git
- Groq API key
- Internet connection for Groq API and web browsing

### 1. Clone the Repository

```bash
git clone https://github.com/amit171ku/PC-AGENT.git
cd PC-AGENT
```

### 2. Create a Virtual Environment

Windows:

```bash
python -m venv venv
venv\Scripts\activate
```

Linux / macOS:

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Install Playwright Chromium

```bash
python -m playwright install chromium
```

### 5. Configure Environment Variables

Create a `.env` file in the project root:

```env
GROQ_API_KEY=your_groq_api_key

# Comma-separated model fallback list
GROQ_MODELS=model_name_1,model_name_2,model_name_3

# Enable or disable automatic fallback
GROQ_AUTO_FALLBACK=on

# Optional workspace location
AGENT_WORKSPACE=

# Optional browser configuration
BROWSER_ALLOW_PRIVATE=false
BROWSER_NO_SANDBOX=false
```

Replace the model names with models currently available to your Groq account.

**Important:** Never upload your `.env` file or expose your API key in GitHub.

### 6. Run the Application

```bash
python main.py
```

Open the application in your browser:

```text
http://127.0.0.1:8000
```

## Model Fallback

PC Agent supports multiple Groq models through the `GROQ_MODELS` configuration.

When a configured model encounters certain errors, the application can switch to another configured model.

The fallback mechanism is designed to handle cases such as:
- Rate-limit responses.
- Unsupported or unavailable models.
- Certain tool-calling errors.

Some temporary network and server errors are retried.

**Note:** Model fallback does not bypass Groq account limits. If all configured models are rate-limited or unavailable, the request can still fail.

## Workspace and File Safety

PC Agent uses a dedicated workspace directory.

By default, the workspace is created under the user's home directory as:

```text
AgentWorkspace/
├── .trash/
├── templates/
└── reports/
```

You can configure a different workspace using `AGENT_WORKSPACE`.

File operations use shared path-validation helpers to restrict access to the configured workspace.

Deleted files are moved to `.trash` rather than being immediately permanently deleted.

## Testing

Run the available automated tests:

```bash
python -m pytest -v
```

The repository also includes a GitHub Actions workflow that installs dependencies and runs pytest.

## Current Limitations

- Groq API access requires an internet connection.
- Browser automation requires Playwright and Chromium.
- Model availability and rate limits depend on your Groq account.
- Browser and file operations are subject to application safety checks, but these controls should not be treated as a complete security boundary.
- The current application is configured for local use and does not provide a complete multi-user authentication system.

## Future Improvements

Potential areas for further development:

- Stronger browser security and SSRF protection.
- More comprehensive automated tests.
- Improved WebSocket connection and confirmation handling.
- Persistent conversation memory.
- More robust error handling and logging.
- User authentication and multi-user support.
- Better task history and execution monitoring.

## Contributing

Contributions, bug reports, and feature suggestions are welcome.

1. Fork the repository.
2. Create a feature branch.
3. Make your changes.
4. Run the available tests.
5. Submit a pull request.

## Author

**Amit Sharma**

- GitHub: https://github.com/amit171ku
- LinkedIn: https://linkedin.com/in/amit-sharma174

---

*PC Agent is a personal AI automation project built to explore LLM tool calling, browser automation, file-system operations, and productivity workflows.*