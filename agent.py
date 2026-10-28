import json
import os
import shutil
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from groq import Groq
from playwright.sync_api import sync_playwright

load_dotenv()  # reads GROQ_API_KEY from .env

# ---------- Config ----------
MODEL = "openai/gpt-oss-120b"
ROOT = Path.home() / "AgentWorkspace"   # the ONLY folder the agent can touch
TRASH = ROOT / ".trash"
LOG = ROOT / "agent_log.txt"
ROOT.mkdir(exist_ok=True)
TRASH.mkdir(exist_ok=True)

client = Groq(api_key=os.environ["GROQ_API_KEY"])

# ---------- Helpers ----------
def safe(path: str) -> Path:
    p = (ROOT / path).resolve()
    if ROOT.resolve() not in p.parents and p != ROOT.resolve():
        raise PermissionError(f"Path outside workspace: {path}")
    return p

def log(msg: str):
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now().isoformat()}  {msg}\n")

def confirm(action: str) -> bool:
    ans = input(f"\n[CONFIRM] {action}  (y/n): ").strip().lower()
    return ans == "y"

# ---------- File tools ----------
def list_dir(path: str = ".") -> str:
    p = safe(path)
    return "\n".join(
        f"{'[DIR] ' if i.is_dir() else ''}{i.name}" for i in sorted(p.iterdir())
    ) or "(empty)"

def read_file(path: str) -> str:
    return safe(path).read_text(encoding="utf-8", errors="replace")[:8000]

def search_files(pattern: str) -> str:
    hits = [str(f.relative_to(ROOT)) for f in ROOT.rglob(pattern) if TRASH not in f.parents]
    return "\n".join(hits[:100]) or "No matches"

def write_file(path: str, content: str) -> str:
    if not confirm(f"Write file {path}"):
        return "User denied."
    p = safe(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(content, encoding="utf-8")
    return f"Written to {p.name}."

def move_file(src: str, dst: str) -> str:
    if not confirm(f"Move {src} -> {dst}"):
        return "User denied."
    d = safe(dst)
    d.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(safe(src)), str(d))
    return "Moved."

def delete_file(path: str) -> str:
    if not confirm(f"Delete {path} (goes to .trash)"):
        return "User denied."
    shutil.move(str(safe(path)), str(TRASH / Path(path).name))
    return "Moved to trash."

# ---------- Playwright Setup & State Tracking ----------
_pw = sync_playwright().start()
_browser = _pw.chromium.launch(headless=False)
_page = _browser.new_page()

# Monitor errors and failed requests for testing/debugging
_console_errors = []
_failed_requests = []
_page.on("console", lambda msg: _console_errors.append(msg.text) if msg.type == "error" else None)
_page.on("requestfailed", lambda req: _failed_requests.append(f"{req.method} {req.url} - {req.failure}"))

# ---------- 1. Browsing & Navigation ----------
def open_url(url: str) -> str:
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    resp = _page.goto(url, wait_until="domcontentloaded", timeout=30000)
    status = resp.status if resp else "unknown"
    return f"Opened: {_page.title()} (Status {status}) at {_page.url}"

def navigate_browser(action: str) -> str:
    """action: 'back', 'forward', or 'reload'"""
    act = action.lower().strip()
    if act == "back":
        _page.go_back(wait_until="domcontentloaded", timeout=15000)
    elif act == "forward":
        _page.go_forward(wait_until="domcontentloaded", timeout=15000)
    elif act == "reload":
        _page.reload(wait_until="domcontentloaded", timeout=15000)
    else:
        return f"Unknown action: '{action}'. Choose 'back', 'forward', or 'reload'."
    return f"Navigated {act}. Current title: {_page.title()} ({_page.url})"

def wait_for(selector: str = "", timeout_ms: str = "2000") -> str:
    ms = int(timeout_ms) if timeout_ms.isdigit() else 2000
    if selector.strip():
        _page.wait_for_selector(selector.strip(), timeout=ms)
        return f"Found selector '{selector}'."
    _page.wait_for_timeout(ms)
    return f"Waited {ms}ms."

# ---------- 2. Automation & Forms ----------
def click(selector: str) -> str:
    if not confirm(f"Click '{selector}' on {_page.url}"):
        return "User denied."
    _page.click(selector, timeout=10000)
    return f"Clicked '{selector}'."

def type_text(selector: str, text: str) -> str:
    if any(w in selector.lower() for w in ("password", "card", "cvv", "secret", "ssn")):
        return "Refused: safety filter prevents typing into sensitive fields."
    if not confirm(f"Type '{text}' into '{selector}'"):
        return "User denied."
    _page.fill(selector, text)
    return f"Typed into '{selector}'."

def select_dropdown(selector: str, value: str) -> str:
    if not confirm(f"Select '{value}' on '{selector}'"):
        return "User denied."
    _page.select_option(selector, value=value, timeout=10000)
    return f"Selected '{value}' in '{selector}'."

def press_key(key: str) -> str:
    _page.keyboard.press(key)
    return f"Pressed '{key}'."

def scroll_page(direction: str = "down", amount: str = "600") -> str:
    px = int(amount) if amount.isdigit() else 600
    delta = px if direction.lower() == "down" else -px
    _page.mouse.wheel(0, delta)
    _page.wait_for_timeout(300)
    return f"Scrolled {direction} by {px}px."

def hover(selector: str) -> str:
    _page.hover(selector, timeout=10000)
    return f"Hovered over '{selector}'."

# ---------- 3. Web Scraping ----------
def get_page_text() -> str:
    return _page.inner_text("body")[:6000]

def extract_links(filter_keyword: str = "") -> str:
    links = _page.eval_on_selector_all(
        "a[href]",
        """els => els.map(a => ({
            text: a.innerText.trim().replace(/\\s+/g, ' '),
            href: a.href
        })).filter(x => x.href && !x.href.startsWith('javascript:'))"""
    )
    if filter_keyword.strip():
        kw = filter_keyword.lower().strip()
        links = [l for l in links if kw in l["text"].lower() or kw in l["href"].lower()]
    links = links[:30]
    if not links:
        return "No matching links found."
    return "\n".join(f"- [{l['text'] or 'link'}]({l['href']})" for l in links)

def extract_table(selector: str = "table") -> str:
    try:
        rows = _page.eval_on_selector(
            selector,
            """t => Array.from(t.querySelectorAll('tr')).map(r =>
                Array.from(r.querySelectorAll('th, td')).map(c => c.innerText.trim().replace(/\\s+/g, ' '))
            )"""
        )
        if not rows:
            return "Table is empty or has no rows."
        headers = rows[0]
        lines = [
            "| " + " | ".join(headers) + " |",
            "| " + " | ".join(["---"] * len(headers)) + " |"
        ]
        for r in rows[1:25]:
            padded = r + [""] * (len(headers) - len(r))
            lines.append("| " + " | ".join(padded[:len(headers)]) + " |")
        if len(rows) > 25:
            lines.append(f"... ({len(rows) - 25} more rows omitted)")
        return "\n".join(lines)
    except Exception as e:
        return f"Error extracting table: {e}"

# ---------- 4. Automated Testing ----------
def test_element(selector: str, state: str = "visible") -> str:
    """state: 'visible', 'hidden', 'attached'"""
    try:
        _page.wait_for_selector(selector, state=state, timeout=5000)
        return f"PASS: Element '{selector}' is {state}."
    except Exception:
        return f"FAIL: Element '{selector}' is NOT {state}."

def test_text(expected_text: str) -> str:
    content = _page.inner_text("body")
    if expected_text in content:
        return f"PASS: Found '{expected_text}' on page."
    return f"FAIL: '{expected_text}' not found on page."

def get_console_errors() -> str:
    errs = []
    if _console_errors:
        errs.append("Console Errors:\n" + "\n".join(f"- {e}" for e in _console_errors[-10:]))
    if _failed_requests:
        errs.append("Failed Requests:\n" + "\n".join(f"- {f}" for f in _failed_requests[-10:]))
    return "\n\n".join(errs) if errs else "PASS: No console errors or failed requests recorded."

# ---------- 5. Screenshots & PDFs ----------
def take_screenshot(path: str = "screenshot.png", selector: str = "") -> str:
    p = safe(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    if selector.strip():
        el = _page.query_selector(selector.strip())
        if not el:
            return f"Selector '{selector}' not found for screenshot."
        el.screenshot(path=str(p))
    else:
        _page.screenshot(path=str(p), full_page=True)
    return f"Screenshot saved to: {p.relative_to(ROOT)}"

def export_pdf(path: str = "page.pdf") -> str:
    p = safe(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    try:
        # Chromium requires headless mode for PDF export; fallback to a temporary headless tab if visible
        _page.pdf(path=str(p), print_background=True)
        return f"PDF saved to: {p.relative_to(ROOT)}"
    except Exception:
        try:
            hl_browser = _pw.chromium.launch(headless=True)
            hl_page = hl_browser.new_page()
            hl_page.goto(_page.url, wait_until="networkidle", timeout=15000)
            hl_page.pdf(path=str(p), print_background=True)
            hl_browser.close()
            return f"PDF saved to: {p.relative_to(ROOT)}"
        except Exception as e:
            return f"Error exporting PDF: {e}"

# ---------- Tool Registry ----------
FUNCS = {
    # File tools
    "list_dir": list_dir, "read_file": read_file, "search_files": search_files,
    "write_file": write_file, "move_file": move_file, "delete_file": delete_file,
    # Browsing & Navigation
    "open_url": open_url, "navigate_browser": navigate_browser, "wait_for": wait_for,
    # Automation
    "click": click, "type_text": type_text, "select_dropdown": select_dropdown,
    "press_key": press_key, "scroll_page": scroll_page, "hover": hover,
    # Scraping
    "get_page_text": get_page_text, "extract_links": extract_links, "extract_table": extract_table,
    # Testing
    "test_element": test_element, "test_text": test_text, "get_console_errors": get_console_errors,
    # Visual & Export
    "take_screenshot": take_screenshot, "export_pdf": export_pdf,
}

def tool(name, desc, props, required):
    return {"type": "function", "function": {
        "name": name, "description": desc,
        "parameters": {"type": "object",
                       "properties": {k: {"type": "string"} for k in props},
                       "required": required}}}

TOOLS = [
    # File Tools
    tool("list_dir", "List files in workspace folder", ["path"], []),
    tool("read_file", "Read a text file in workspace", ["path"], ["path"]),
    tool("search_files", "Find files by glob pattern (e.g. *.png)", ["pattern"], ["pattern"]),
    tool("write_file", "Create or overwrite a file in workspace", ["path", "content"], ["path", "content"]),
    tool("move_file", "Move or rename a file in workspace", ["src", "dst"], ["src", "dst"]),
    tool("delete_file", "Delete a file (moves to .trash)", ["path"], ["path"]),

    # 1. Browsing & Navigation
    tool("open_url", "Open a website URL", ["url"], ["url"]),
    tool("navigate_browser", "Navigate browser history: 'back', 'forward', or 'reload'", ["action"], ["action"]),
    tool("wait_for", "Wait for a selector to appear or pause (timeout_ms)", ["selector", "timeout_ms"], []),

    # 2. Automation
    tool("click", "Click an element by CSS selector", ["selector"], ["selector"]),
    tool("type_text", "Type into an input field by CSS selector", ["selector", "text"], ["selector", "text"]),
    tool("select_dropdown", "Select option from a <select> by value", ["selector", "value"], ["selector", "value"]),
    tool("press_key", "Press a key like 'Enter', 'Tab', 'Escape'", ["key"], ["key"]),
    tool("scroll_page", "Scroll the page ('up' or 'down', amount in px)", ["direction", "amount"], []),
    tool("hover", "Hover over an element by CSS selector", ["selector"], ["selector"]),

    # 3. Scraping
    tool("get_page_text", "Get visible text content of current page", [], []),
    tool("extract_links", "Extract page links (optionally filtered by keyword)", ["filter_keyword"], []),
    tool("extract_table", "Extract an HTML table into a Markdown table", ["selector"], []),

    # 4. Testing & Verification
    tool("test_element", "Assert an element state ('visible' or 'hidden')", ["selector", "state"], ["selector"]),
    tool("test_text", "Assert specific text is present on the page", ["expected_text"], ["expected_text"]),
    tool("get_console_errors", "Check for JavaScript console errors or network failures", [], []),

    # 5. Screenshots & PDFs
    tool("take_screenshot", "Capture page or element screenshot to workspace (e.g. shot.png)", ["path", "selector"], ["path"]),
    tool("export_pdf", "Export current page as PDF to workspace (e.g. page.pdf)", ["path"], ["path"]),
]

SYSTEM = (
    f"You are a helpful PC assistant. Workspace: {ROOT} (use relative paths for all file operations). "
    "You have real Playwright browser tools:\n"
    "- Navigation: open_url, navigate_browser, wait_for.\n"
    "- Automation: click, type_text, select_dropdown, press_key, scroll_page, hover.\n"
    "- Scraping: get_page_text, extract_links, extract_table.\n"
    "- Testing: test_element, test_text, get_console_errors.\n"
    "- Visual/Export: take_screenshot, export_pdf.\n"
    "Always verify actions using tools. Keep replies concise and never attempt to bypass user confirmations."
)

# ---------- Agent loop ----------
def run(messages):
    for _ in range(15):
        resp = client.chat.completions.create(
            model=MODEL, messages=messages, tools=TOOLS, tool_choice="auto"
        )
        msg = resp.choices[0].message
        messages.append(msg)
        if not msg.tool_calls:
            return msg.content
        for call in msg.tool_calls:
            name = call.function.name
            try:
                args = json.loads(call.function.arguments or "{}")
            except Exception:
                args = {}
            try:
                result = FUNCS[name](**args)
            except Exception as e:
                result = f"Error: {e}"
            log(f"{name}({args}) -> {str(result)[:200]}")
            print(f"  [tool] {name}({args})")
            messages.append({"role": "tool", "tool_call_id": call.id, "content": str(result)})
    return "Stopped: maximum steps reached."

if __name__ == "__main__":
    print(f"Agent ready. Workspace: {ROOT}\nType 'exit' to quit.")
    history = [{"role": "system", "content": SYSTEM}]
    try:
        while True:
            q = input("\nYou: ").strip()
            if q.lower() in ("exit", "quit"):
                break
            if not q:
                continue
            history.append({"role": "user", "content": q})
            print("\nAgent:", run(history))
    finally:
        _browser.close()
        _pw.stop()