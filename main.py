import asyncio
import json
import os
import queue
import re
import shutil
import threading
import time
import uuid
from contextlib import asynccontextmanager
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus

import uvicorn
from dotenv import load_dotenv

load_dotenv(override=True)   # first, so settings such as AGENT_WORKSPACE in .env apply to every module

from fastapi import (FastAPI, File, HTTPException, Request, UploadFile,  # noqa: E402
                     WebSocket, WebSocketDisconnect)
from fastapi.responses import FileResponse  # noqa: E402
from fastapi.staticfiles import StaticFiles  # noqa: E402
from groq import APIStatusError, Groq, RateLimitError  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

import common  # noqa: E402
import design_tools  # noqa: E402  (restyle existing Word / PowerPoint files)
import doc_tools  # noqa: E402  (read_document, export_pdf)
import excel_tools  # noqa: E402
import ppt_tools  # noqa: E402
import word_tools  # noqa: E402
from common import LOG, ROOT, TRASH, safe, tool  # noqa: E402

# ---------- Config ----------
BASE = Path(__file__).parent
PORT = 8000
ALLOWED_ORIGINS = {f"http://127.0.0.1:{PORT}", f"http://localhost:{PORT}"}
MAX_STEPS = 40
UPLOAD_EXT = {".docx", ".xlsx", ".pptx", ".pdf", ".txt", ".md", ".csv",
              ".png", ".jpg", ".jpeg"}

client = Groq(api_key=os.environ["GROQ_API_KEY"])


CHAT_SKIP = ("guard", "whisper", "tts", "orpheus", "embed", "allam")
_blocked = {}                 # model -> time when it can be used again (rate limited)
_dead = {}                    # model -> reason it cannot be used this session
_current = {"name": ""}       # model currently answering


def _wait_secs(err):
    """Seconds Groq asks us to wait, parsed from the error text (default 20)."""
    msg = str(err)
    m = re.search(r"try again in\s+([0-9hms.]+)", msg)
    if m:
        total = 0.0
        for val, unit in re.findall(r"(\d+(?:\.\d+)?)(h|m(?!s)|s)", m.group(1)):
            total += float(val) * {"h": 3600, "m": 60, "s": 1}[unit]
        if total:
            return total
    return 20.0


def build_model_list():
    """Priority order: GROQ_MODELS from .env (comma separated; GROQ_MODEL also works), then
    automatic extras unless GROQ_AUTO_FALLBACK=off. Groq limits are per model, so moving to
    the next model when one is limited keeps the agent working."""
    available = [m.id for m in client.models.list().data]
    chat = [m for m in available if not any(x in m.lower() for x in CHAT_SKIP)]
    raw = os.getenv("GROQ_MODELS", "") or os.getenv("GROQ_MODEL", "")
    wanted = [m.strip() for m in raw.split(",") if m.strip()]
    for m in wanted:
        if m not in available:
            print(f"[models] WARNING: '{m}' is not available on your Groq account and will be skipped")
    order = [m for m in wanted if m in available]
    if os.getenv("GROQ_AUTO_FALLBACK", "on").strip().lower() != "off":
        order += [m for m in ("openai/gpt-oss-120b", "openai/gpt-oss-20b") if m in available]
        order += [m for m in chat if m.startswith("qwen")]
    seen, out = set(), []
    for m in order:
        if m not in seen:
            seen.add(m)
            out.append(m)
    return out or chat


MODELS = build_model_list()
print("Models in priority order:", MODELS)


def verify_models():
    """Check at startup that at least one model accepts tool calling."""
    test_tool = [{"type": "function", "function": {
        "name": "ping", "description": "test",
        "parameters": {"type": "object", "properties": {}, "required": []}}}]
    for m in MODELS:
        try:
            client.chat.completions.create(
                model=m, messages=[{"role": "user", "content": "hi"}],
                tools=test_tool, max_tokens=20)
            print("Ready:", m)
            if not _current["name"]:
                _current["name"] = m
        except RateLimitError as e:
            _blocked[m] = time.time() + _wait_secs(e)
            print("Rate limited right now (will retry later):", m)
            if not _current["name"]:
                _current["name"] = m
        except Exception as e:
            _dead[m] = str(e)[:80]
            print("Skipped", m, "->", str(e)[:90])
    if not _current["name"]:
        raise SystemExit("No model with tool calling found. Check GROQ_MODELS in .env.")


verify_models()
MODEL = _current["name"]

# ---------- Messaging between worker thread and browser UI ----------
cmd_q = queue.Queue()
clients = set()
pending = {}          # confirm id -> {"event": Event, "ok": bool}
answers = {}          # question id -> {"event": Event, "text": str}
loop = None


async def broadcast(event: dict):
    for ws in list(clients):
        try:
            await ws.send_text(json.dumps(event))
        except Exception:
            clients.discard(ws)


def emit(event: dict):
    if loop is not None:
        asyncio.run_coroutine_threadsafe(broadcast(event), loop)


def confirm(action: str) -> bool:
    cid = uuid.uuid4().hex
    box = {"event": threading.Event(), "ok": False}
    pending[cid] = box
    emit({"type": "confirm", "id": cid, "text": action})
    box["event"].wait(timeout=120)      # no answer in 2 minutes = deny
    pending.pop(cid, None)
    return box["ok"]


common.set_confirm_handler(confirm)   # Word/Excel/PowerPoint modules ask through the same browser prompt


# ---------- Helpers ----------
def log(msg: str):
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now().isoformat()}  {msg}\n")


# ---------- File tools ----------
def list_dir(path: str = ".") -> str:
    p = safe(path)
    return "\n".join(
        f"{'[DIR] ' if i.is_dir() else ''}{i.name}" for i in sorted(p.iterdir())
    ) or "(empty)"


def read_file(path: str) -> str:
    return safe(path).read_text(encoding="utf-8", errors="replace")[:3000]


def search_files(pattern: str) -> str:
    hits = [str(f.relative_to(ROOT)) for f in ROOT.rglob(pattern) if TRASH not in f.parents]
    return "\n".join(hits[:100]) or "No matches"


def write_file(path: str, content: str) -> str:
    if not confirm(f"Write file: {path}"):
        return "User denied."
    p = safe(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(content, encoding="utf-8")
    return "Written."


def move_file(src: str, dst: str) -> str:
    if not confirm(f"Move: {src}  ->  {dst}"):
        return "User denied."
    d = safe(dst)
    d.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(safe(src)), str(d))
    return "Moved."


def delete_file(path: str) -> str:
    if not confirm(f"Delete (moves to .trash): {path}"):
        return "User denied."
    shutil.move(str(safe(path)), str(TRASH / Path(path).name))
    return "Moved to trash."


# ---------- Browser tools (created lazily inside the worker thread) ----------
_b = {}


def page():
    br = _b.get("browser")
    pg = _b.get("page")
    try:
        if pg is not None and br is not None and br.is_connected() and not pg.is_closed():
            return pg
    except Exception:
        pass
    try:
        if br is not None:
            br.close()
    except Exception:
        pass
    try:
        if "pw" in _b:
            _b["pw"].stop()
    except Exception:
        pass
    _b.clear()
    _b["pw"] = sync_playwright().start()
    _b["browser"] = _b["pw"].chromium.launch(headless=False)
    _b["page"] = _b["browser"].new_page()
    return _b["page"]


def search_web(query: str) -> str:
    page().goto("https://duckduckgo.com/html/?q=" + quote_plus(query), timeout=30000)
    return page().inner_text("body")[:3000]


def open_url(url: str) -> str:
    page().goto(url, timeout=30000)
    return f"Opened {page().title()}"


def get_page_text() -> str:
    return page().inner_text("body")[:3000]


def click(selector: str) -> str:
    if not confirm(f"Click: {selector}  on {page().url}"):
        return "User denied."
    page().click(selector, timeout=10000)
    return "Clicked."


def type_text(selector: str, text: str) -> str:
    if any(w in selector.lower() for w in ("password", "card", "cvv")):
        return "Refused: the agent does not type into sensitive fields."
    if not confirm(f"Type: {text}  into {selector}"):
        return "User denied."
    page().fill(selector, text)
    return "Typed."


def close_browser():
    try:
        if "browser" in _b:
            _b["browser"].close()
        if "pw" in _b:
            _b["pw"].stop()
    except Exception:
        pass


# ---------- Ask the user a question (waits for an answer in the UI) ----------
def ask_user(question: str) -> str:
    qid = uuid.uuid4().hex
    box = {"event": threading.Event(), "text": ""}
    answers[qid] = box
    emit({"type": "ask", "id": qid, "text": question})
    box["event"].wait(timeout=900)
    answers.pop(qid, None)
    return box["text"].strip() or (
        "(No answer given. Proceed with sensible defaults and mark unknown details as TBD.)")


# ---------- Tool registry ----------
CORE_FUNCS = {f.__name__: f for f in [
    list_dir, read_file, search_files, write_file, move_file, delete_file,
    search_web, open_url, get_page_text, click, type_text,
]}

ASK_SCHEMA = tool("ask_user", "Ask the user clarifying questions and wait. ONE call, short numbered list.",
                  {"question": "Numbered questions"}, ["question"])

# Each module lists (group, schema, function). Groups decide which tools are sent to the model.
REGISTRY = ([("base", ASK_SCHEMA, ask_user)]
            + doc_tools.REGISTRY
            + word_tools.REGISTRY
            + excel_tools.REGISTRY
            + ppt_tools.REGISTRY
            + design_tools.REGISTRY)
FUNCS = {**CORE_FUNCS, **{schema["function"]["name"]: fn for _, schema, fn in REGISTRY}}
TOOLS_OFFICE = [schema for _, schema, _ in REGISTRY]
TOOL_GROUP = {schema["function"]["name"]: group for group, schema, _ in REGISTRY}

TOOLS_CORE = [
    tool("list_dir", "List files in a workspace folder", ["path"], []),
    tool("read_file", "Read a small text file", ["path"], ["path"]),
    tool("search_files", "Find files by glob pattern", ["pattern"], ["pattern"]),
    tool("write_file", "Create or overwrite a text file", ["path", "content"], ["path", "content"]),
    tool("move_file", "Move or rename a file", ["src", "dst"], ["src", "dst"]),
    tool("delete_file", "Delete a file (moved to trash)", ["path"], ["path"]),
    tool("search_web", "Search the web, returns result text", ["query"], ["query"]),
    tool("open_url", "Open a URL in the browser", ["url"], ["url"]),
    tool("get_page_text", "Get visible text of the current page", [], []),
    tool("click", "Click an element by CSS selector", ["selector"], ["selector"]),
    tool("type_text", "Type into an input by CSS selector", ["selector", "text"], ["selector", "text"]),
]

GROUP_RE = {
    "excel": re.compile(r"excel|xlsx|spreadsheet|sheet|workbook|budget|tracker|cell|column|conditional|highlight|kpi|dashboard", re.I),
    "ppt": re.compile(r"powerpoint|pptx|presentation|slides?|deck", re.I),
    "word": re.compile(r"word|docx|report|proposal|memo|letter|document|brief|restyle|redesign", re.I),
    "pdf": re.compile(r"pdf", re.I),
}


def build_tools(groups):
    keep = set(groups) | {"base"}
    return TOOLS_CORE + [t for t in TOOLS_OFFICE if TOOL_GROUP[t["function"]["name"]] in keep]


OFFICE_RE = re.compile(
    r"report|proposal|word|docx|excel|xlsx|spreadsheet|sheet|workbook|budget|tracker|format|cell|column|highlight|conditional|redesign|design|kpi|portfolio|premium|dashboard|"
    r"powerpoint|pptx|presentation|slides?|deck|template|pdf|document|memo|letter|invoice|brief",
    re.I)

BASE_PROMPT = (
    f"You are a PC assistant. Workspace folder: {ROOT} (always use relative paths). "
    "Folders: templates/ holds the user's samples and templates; reports/ is where finished files go. "
    "You have a real browser (search_web, open_url, get_page_text). For ANY request about searching, "
    "websites, news, prices or current information you MUST call search_web or open_url and never answer from memory. "
    "Keep chat replies short and never bypass confirmations."
)

OFFICE_PROMPT = (
    "\nDOCUMENT WORK: pick the format that fits: Word (.docx) for reports, proposals, letters and memos; "
    "Excel (.xlsx) for budgets, trackers, schedules, data and calculations; PowerPoint (.pptx) for presentations; "
    "export_pdf only when a PDF is wanted. You may combine formats. If the format is unclear, ask.\n"
    "WORKFLOW: (1) list_dir templates/ and read_document any relevant sample or template; follow its structure, "
    "headings, numbering, columns and tone exactly. (2) If key facts are missing (project name, client, dates, "
    "budget, results, audience, length), call ask_user ONCE with a short numbered list. Never invent facts, figures "
    "or dates; use 'TBD'. (3) Build in reports/. Word: report_start (pass template= when a .docx template exists), "
    "then report_add_content with markdown, several sections per call (about 300 to 600 words), and report_add_chart "
    "when numbers deserve a chart. Excel: excel_start, excel_add_sheet per table, excel_set_cells for totals "
    "(like B7==SUM(B2:B6)), excel_add_chart. PowerPoint: ppt_start, ONE ppt_add_slides call with 3 to 6 short bullets "
    "per slide, plus ppt_add_table_slide or ppt_add_chart_slide where useful. (4) Finish with the file path(s), "
    "a 2 to 3 line summary, and remind the user to review the files before sharing.\n"
    "FULL RESTYLE of an existing Word file: call word_design_report ONCE; of an existing PowerPoint: call ppt_design_deck ONCE (pass an output name such as reports/Name_Designed.docx or .pptx). Do not hand-format. Reply with only the output file and a short list of changes.\n"
    "FULL REDESIGN of a sheet (premium, portfolio, dashboard, KPI, polished): call excel_design_report ONCE with the exact title and subtitle the user gave and the requested output file name. Do not hand-format cell by cell. Reply with only the output file and a short list of changes.\n"
    "EDITING AN EXISTING EXCEL FILE: first excel_read it to see the real columns and rows, then use excel_format "
    "(fonts, fills, borders, number formats), excel_conditional_format (color scales, data bars, highlight rules, "
    "zebra stripes) and excel_layout (freeze header, filter, column widths) with exact ranges. Apply the changes; "
    "do not say a tool is missing."
)


# ---------- Token-saving helpers ----------
def compact(messages, tail=6, tool_tail=10, limit=1200):
    """Shrink old, bulky tool arguments/results so long builds stay under rate limits."""
    for m in messages[:max(len(messages) - tail, 0)]:
        if m.get("role") == "assistant" and m.get("tool_calls"):
            for c in m["tool_calls"]:
                if c["function"]["name"].startswith(("report_add", "ppt_add", "excel_add", "write_file")):
                    try:
                        a = json.loads(c["function"]["arguments"])
                    except Exception:
                        continue
                    short = {k: (v[:60] + "...") if isinstance(v, str) and len(v) > 60 else v
                             for k, v in a.items()}
                    c["function"]["arguments"] = json.dumps(short)
    for m in messages[:max(len(messages) - tool_tail, 0)]:
        if m.get("role") == "tool" and len(m.get("content", "")) > limit:
            m["content"] = m["content"][:limit] + " [trimmed]"


def trim_history(history, keep=24):
    """Keep the system prompt and recent messages, starting on a user message
    so tool results are never separated from their tool calls."""
    if len(history) <= keep + 1:
        return history
    tail = history[-keep:]
    for i, m in enumerate(tail):
        if m.get("role") == "user":
            return [history[0]] + tail[i:]
    return [history[0], history[-1]]


def _switch(model, why):
    """Tell the UI when the active model changes."""
    old = _current["name"]
    if old != model:
        _current["name"] = model
        print(f"[models] {old} -> {model}: {why}")
        emit({"type": "model", "name": model, "note": f"Switched to {model} ({why})"})


def chat_call(messages, tools):
    shrunk, skip = False, set()
    reason = "preferred model is available again"
    for _ in range(14):
        now = time.time()
        live = [m for m in MODELS if m not in _dead and m not in skip]
        usable = [m for m in live if _blocked.get(m, 0) <= now]
        if not usable:
            waits = [_blocked[m] - now for m in live if _blocked.get(m, 0) > now]
            if not waits:
                raise RuntimeError("No Groq model could answer. Check GROQ_MODELS in .env.")
            wait = min(waits)
            if wait > 120:
                raise RuntimeError(
                    f"Every model hit its Groq limit. Try again in about {int(wait // 60) + 1} min "
                    "(daily token limits can take longer). Click New chat to use fewer tokens.")
            emit({"type": "tool", "text": "rate_limit()"})
            time.sleep(wait + 1)
            continue
        model = usable[0]
        _switch(model, reason)
        extra = {"max_completion_tokens": 3000}
        if model.startswith("openai/gpt-oss"):
            extra["reasoning_effort"] = "low"
        try:
            return client.chat.completions.create(
                model=model, messages=messages, tools=tools, tool_choice="auto",
                extra_body=extra)
        except RateLimitError as e:
            secs = _wait_secs(e)
            _blocked[model] = time.time() + secs
            reason = f"{model} hit its limit"
            print(f"[limit] {model} blocked for {secs:.0f}s")
            continue
        except APIStatusError as e:
            msg = str(e).lower()
            if e.status_code == 413 and not shrunk:
                compact(messages, tail=0, tool_tail=0, limit=400)
                shrunk = True
                continue
            if e.status_code == 413:
                raise RuntimeError("This request is too large for Groq's free per-minute limit. "
                                   "Click New chat and ask for a smaller piece at a time.")
            if "tool_use_failed" in msg or "failed to call a function" in msg:
                skip.add(model)
                reason = f"{model} could not format a tool call"
                continue
            if e.status_code in (400, 403, 404) and any(x in msg for x in (
                    "tool calling", "not supported", "does not exist", "model_not_found",
                    "decommission", "permission", "terms")):
                _dead[model] = str(e)[:80]
                reason = f"{model} is unavailable"
                continue
            raise
    raise RuntimeError("Groq rate limit reached on all models. Wait a minute and try again.")


# ---------- Agent loop ----------
def run(messages, tools):
    for _ in range(MAX_STEPS):
        compact(messages)
        resp = chat_call(messages, tools)
        msg = resp.choices[0].message
        entry = {"role": "assistant", "content": msg.content}
        if msg.tool_calls:
            entry["tool_calls"] = [
                {"id": c.id, "type": "function",
                 "function": {"name": c.function.name,
                              "arguments": c.function.arguments or "{}"}}
                for c in msg.tool_calls]
        messages.append(entry)
        if not msg.tool_calls:
            return msg.content
        for call in msg.tool_calls:
            name = call.function.name
            try:
                args = json.loads(call.function.arguments or "{}")
            except json.JSONDecodeError:
                args = {}
            emit({"type": "tool", "text": f"{name}({str(args)[:200]})"})
            try:
                result = FUNCS[name](**args)
            except Exception as e:
                result = f"Error: {e}"
            log(f"{name}({str(args)[:200]}) -> {str(result)[:200]}")
            print(f"[tool] {name} -> {str(result)[:150]}")
            messages.append({"role": "tool", "tool_call_id": call.id, "content": str(result)})
    return "Stopped: too many steps. Ask me to continue."


def worker():
    office = False
    groups = set()
    history = [{"role": "system", "content": BASE_PROMPT}]
    while True:
        text = cmd_q.get()
        if text is None:
            close_browser()
            break
        if text == "__RESET__":
            office = False
            groups = set()
            history = [{"role": "system", "content": BASE_PROMPT}]
            continue
        if OFFICE_RE.search(text):
            office = True
            found = {g for g, rx in GROUP_RE.items() if rx.search(text)}
            groups |= found if found else ({"word", "excel", "ppt"} if not groups else set())
        history.append({"role": "user", "content": text})
        history = trim_history(history)
        history[0]["content"] = BASE_PROMPT + (OFFICE_PROMPT if office else "")
        tools = build_tools(groups) if office else TOOLS_CORE
        try:
            answer = run(history, tools)
        except Exception as e:
            answer = f"Error: {e}"
        emit({"type": "agent", "text": answer or ""})
        emit({"type": "done"})


# ---------- FastAPI ----------
@asynccontextmanager
async def lifespan(app):
    global loop
    loop = asyncio.get_running_loop()
    threading.Thread(target=worker, daemon=True).start()
    yield
    cmd_q.put(None)


app = FastAPI(lifespan=lifespan)
app.mount("/static", StaticFiles(directory=BASE / "static"), name="static")


@app.get("/")
async def index():
    return FileResponse(BASE / "static" / "index.html")


@app.get("/info")
async def info():
    return {"model": _current["name"], "workspace": str(ROOT)}


@app.get("/models")
async def models():
    now = time.time()
    out = []
    for m in MODELS:
        if m in _dead:
            out.append({"name": m, "state": "unavailable"})
        elif _blocked.get(m, 0) > now:
            out.append({"name": m, "state": "limited", "wait": int(_blocked[m] - now) + 1})
        else:
            out.append({"name": m, "state": "ready"})
    return {"current": _current["name"], "models": out}


@app.get("/files")
async def files():
    items = []
    for p in sorted(ROOT.rglob("*"), key=lambda x: str(x).lower()):
        rel = p.relative_to(ROOT)
        if any(part.startswith(".") for part in rel.parts) or len(rel.parts) > 3:
            continue
        items.append({
            "path": rel.as_posix(),
            "name": p.name,
            "dir": p.is_dir(),
            "size": p.stat().st_size if p.is_file() else None,
        })
        if len(items) >= 300:
            break
    return items


@app.get("/download")
async def download(path: str):
    try:
        p = safe(path)
    except PermissionError:
        raise HTTPException(status_code=403, detail="Not allowed")
    if not p.is_file():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(p, filename=p.name)


@app.post("/upload")
async def upload(request: Request, file: UploadFile = File(...)):
    origin = request.headers.get("origin")
    if origin is not None and origin not in ALLOWED_ORIGINS:
        raise HTTPException(status_code=403, detail="Forbidden origin")
    name = re.sub(r"[^\w.\- ]", "_", Path(file.filename or "upload").name)
    if Path(name).suffix.lower() not in UPLOAD_EXT:
        raise HTTPException(status_code=400,
                            detail="Allowed types: " + ", ".join(sorted(UPLOAD_EXT)))
    data = await file.read()
    if len(data) > 15 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="File is larger than 15 MB")
    (ROOT / "templates" / name).write_bytes(data)
    return {"saved": f"templates/{name}"}


@app.websocket("/ws")
async def ws_endpoint(ws: WebSocket):
    # Block other websites from talking to the agent
    if ws.headers.get("origin") not in ALLOWED_ORIGINS:
        await ws.close(code=1008)
        return
    await ws.accept()
    clients.add(ws)
    try:
        while True:
            data = json.loads(await ws.receive_text())
            t = data.get("type")
            if t == "message" and data.get("text", "").strip():
                cmd_q.put(data["text"].strip())
            elif t == "reset":
                cmd_q.put("__RESET__")
            elif t == "confirm":
                box = pending.get(data.get("id"))
                if box:
                    box["ok"] = bool(data.get("ok"))
                    box["event"].set()
            elif t == "answer":
                box = answers.get(data.get("id"))
                if box:
                    box["text"] = str(data.get("text", ""))[:4000]
                    box["event"].set()
    except WebSocketDisconnect:
        clients.discard(ws)


if __name__ == "__main__":
    # 127.0.0.1 = only this PC can reach it. Do not change to 0.0.0.0.
    uvicorn.run(app, host="127.0.0.1", port=PORT)