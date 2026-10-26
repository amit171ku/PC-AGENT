import json
import os
import shutil
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from groq import Groq
from playwright.sync_api import sync_playwright

load_dotenv()  # reads GROQ_API_KEY from the .env file

# ---------- Config ----------
MODEL = "openai/gpt-oss-120b"  # check console.groq.com/docs/models for current tool-use models
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
    return "Written."

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

# ---------- Browser tools ----------
_pw = sync_playwright().start()
_browser = _pw.chromium.launch(headless=False)
_page = _browser.new_page()

def open_url(url: str) -> str:
    _page.goto(url, timeout=30000)
    return f"Opened {_page.title()}"

def get_page_text() -> str:
    return _page.inner_text("body")[:8000]

def click(selector: str) -> str:
    if not confirm(f"Click '{selector}' on {_page.url}"):
        return "User denied."
    _page.click(selector, timeout=10000)
    return "Clicked."

def type_text(selector: str, text: str) -> str:
    if any(w in selector.lower() for w in ("password", "card", "cvv")):
        return "Refused: the agent does not type into sensitive fields."
    if not confirm(f"Type '{text}' into '{selector}'"):
        return "User denied."
    _page.fill(selector, text)
    return "Typed."

# ---------- Tool registry ----------
FUNCS = {f.__name__: f for f in [
    list_dir, read_file, search_files, write_file, move_file, delete_file,
    open_url, get_page_text, click, type_text,
]}

def tool(name, desc, props, required):
    return {"type": "function", "function": {
        "name": name, "description": desc,
        "parameters": {"type": "object",
                       "properties": {k: {"type": "string"} for k in props},
                       "required": required}}}

TOOLS = [
    tool("list_dir", "List files in a folder inside the workspace", ["path"], []),
    tool("read_file", "Read a text file", ["path"], ["path"]),
    tool("search_files", "Find files by glob pattern, e.g. *.pdf", ["pattern"], ["pattern"]),
    tool("write_file", "Create or overwrite a text file", ["path", "content"], ["path", "content"]),
    tool("move_file", "Move or rename a file", ["src", "dst"], ["src", "dst"]),
    tool("delete_file", "Delete a file (moved to trash)", ["path"], ["path"]),
    tool("open_url", "Open a URL in the browser", ["url"], ["url"]),
    tool("get_page_text", "Get visible text of the current page", [], []),
    tool("click", "Click an element by CSS selector", ["selector"], ["selector"]),
    tool("type_text", "Type into an input by CSS selector", ["selector", "text"], ["selector", "text"]),
]

SYSTEM = (
    f"You are a helpful PC assistant. You can manage files inside {ROOT} "
    "(use relative paths) and browse the web. Use tools to act, keep answers short, "
    "and never attempt to bypass confirmations."
)

# ---------- Agent loop ----------
def run(messages):
    for _ in range(12):
        resp = client.chat.completions.create(
            model=MODEL, messages=messages, tools=TOOLS, tool_choice="auto"
        )
        msg = resp.choices[0].message
        messages.append(msg)
        if not msg.tool_calls:
            return msg.content
        for call in msg.tool_calls:
            name = call.function.name
            args = json.loads(call.function.arguments or "{}")
            try:
                result = FUNCS[name](**args)
            except Exception as e:
                result = f"Error: {e}"
            log(f"{name}({args}) -> {str(result)[:200]}")
            print(f"  [tool] {name}({args})")
            messages.append({"role": "tool", "tool_call_id": call.id, "content": str(result)})
    return "Stopped: too many steps."

if __name__ == "__main__":
    print(f"Agent ready. Workspace: {ROOT}\nType 'exit' to quit.")
    history = [{"role": "system", "content": SYSTEM}]
    while True:
        q = input("\nYou: ").strip()
        if q.lower() in ("exit", "quit"):
            break
        history.append({"role": "user", "content": q})
        print("Agent:", run(history))
    _browser.close()
    _pw.stop()