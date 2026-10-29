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
- Automatically try eligible fallback models when configured.
- Reset conversations when starting a new task.

### File Management
- List, read, search, and manage files using registered tools.
- Create, move, rename, and delete supported files.
- Use workspace path validation to help prevent unauthorized file access.
- Keep deleted files in a trash directory where supported.
- Maintain an action log.

### Browser Automation
- Search the web and open webpages.
- Extract webpage text, links, and tables.
- Interact with supported webpage elements using Playwright.
- Check elements and text on a page.
- Inspect browser console errors.
- Capture screenshots and export pages to PDF.
- Use confirmation handling for supported sensitive actions.

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

---

## Project Structure

```text
PC-AGENT/
├── main.py
├── common.py
├── browser_tools.py
├── design_tools.py
├── doc_tools.py
├── word_tools.py
├── word_design.py
├── excel_tools.py
├── excel_design.py
├── ppt_tools.py
├── ppt_design.py
├── requirements.txt
├── .gitignore
├── README.md
└── static/
    ├── index.html
    ├── style.css
    └── app.js
```

### Important files

| File | Responsibility |
|---|---|
| `main.py` | FastAPI application, agent loop, model setup, and tool registration |
| `common.py` | Shared utilities, workspace paths, logging, and confirmation helpers |
| `browser_tools.py` | Browser automation tools |
| `word_tools.py` | Word document operations |
| `word_design.py` | Word styling and design helpers |
| `excel_tools.py` | Excel workbook operations |
| `excel_design.py` | Excel styling and design helpers |
| `ppt_tools.py` | PowerPoint operations |
| `ppt_design.py` | PowerPoint styling and design helpers |
| `doc_tools.py` | Additional document utilities |
| `design_tools.py` | Shared design-related functionality |
| `static/` | Web interface assets |
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

Add your Groq API key and model configuration:

```dotenv
GROQ_API_KEY=your_groq_api_key
GROQ_MODELS=your_primary_model,your_fallback_model
GROQ_AUTO_FALLBACK=on
```

Replace the example values with valid model IDs available to your Groq account.

The application supports:

- `GROQ_API_KEY` — authenticates requests to the Groq API.
- `GROQ_MODELS` — a comma-separated list of preferred models in priority order.
- `GROQ_MODEL` — an alternative single-model configuration when `GROQ_MODELS` is not set.
- `GROQ_AUTO_FALLBACK` — controls whether the application adds its built-in fallback candidates. Set it to `off` to disable that behavior.

Important: The current code does not use `GROQ_MODEL_1`, `GROQ_MODEL_2`, or `GROQ_MODEL_3`. Use the variable names shown above.

Never commit your `.env` file or expose your API key publicly.

### Step 5: Run the application

```bash
python main.py
```

After startup, open:

http://127.0.0.1:8000

Keep the terminal running while you use the application.

To stop the server, press `Ctrl+C`.

---

## Example Use Cases

### File Management

Try requests such as:

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

PC Agent builds a model priority list using the configuration in `.env` and the models available through your Groq account.

When automatic fallback is enabled, the application can add eligible fallback candidates from its implementation.

Model fallback may help when a particular model is unavailable or rate-limited. However, switching models does not guarantee uninterrupted service or bypass Groq account limits.

For best results:

1. Use model IDs that are currently available to your account.
2. Confirm that the selected models support the required tool-calling interface.
3. Check the startup logs for the model priority list and any skipped models.
4. Monitor API limits and usage through the Groq Console.

Groq resources: https://console.groq.com/docs/models

---

## Security Considerations

PC Agent can interact with files and websites, so security matters.

- Keep API keys and credentials in local environment variables.
- Use a dedicated workspace for testing.
- Review confirmation prompts before approving sensitive actions.
- Treat webpage content and model-generated instructions as untrusted input.
- Review generated documents and browser actions before relying on them.
- Keep the server bound to localhost unless you have deliberately reviewed the implications of exposing it.
- Test file operations with disposable files before using important data.

Path validation and confirmation prompts are useful safeguards, but they do not guarantee complete security. Review the code and permissions before using the agent with sensitive information.

---

## Troubleshooting

| Problem | Possible solution |
|---|---|
| Missing Groq API key | Check that `.env` is beside `main.py` and contains `GROQ_API_KEY`. |
| Model is unavailable | Verify the model ID and check availability in your Groq account. |
| Rate limit or HTTP 429 | Wait for the limit to reset, reduce request frequency, or configure another eligible model. |
| Tool-calling error | Use a model that supports the required tool-calling interface. |
| `ModuleNotFoundError` | Activate your virtual environment and reinstall `requirements.txt`. |
| Chromium fails to launch | Run `python -m playwright install chromium`. |
| Port 8000 is already in use | Stop the other process or change the port configuration in `main.py`. |
| Web interface does not load | Check the terminal logs and confirm that the `static/` directory exists. |
| Office file generation fails | Check input files, dependencies, permissions, and tool registration. |

---

## Development and Customization

### Add a new tool

1. Implement the Python function in the appropriate module.
2. Define its tool schema.
3. Register the function and schema in the agent.
4. Test the function independently before integrating it.
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
