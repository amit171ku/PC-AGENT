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

import uvicorn
from dotenv import load_dotenv

load_dotenv(override=True)   # settings in .env apply to every module

from fastapi import (FastAPI, File, HTTPException, Request, UploadFile,
                     WebSocket, WebSocketDisconnect)
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from groq import APIStatusError, Groq, RateLimitError

import browser_tools
import common
import design_tools
import doc_tools
import excel_tools
import ppt_tools
import word_tools
from common import LOG, ROOT, TRASH, safe, tool

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
    """Seconds Groq asks us to wait, parsed from error text."""
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
    box["event"].wait(timeout=120)      # 2 minutes timeout = deny
    pending.pop(cid, None)
    return box["ok"]


# One approval dialog for every module
common.set_confirm_handler(confirm)
browser_tools.set_confirm_handler(confirm)


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
    return safe(path).read_text(encoding="utf-8", errors="replace")[:4000]


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


# ---------- Ask user question ----------
def ask_user(question: str) -> str:
    qid = uuid.uuid4().hex
    box = {"event": threading.Event(), "text": ""}
    answers[qid] = box
    emit({"type": "ask", "id": qid, "text": question})
    box["event"].wait(timeout=900)
    answers.pop(qid, None)
    return box["text"].strip() or "(No answer given. Proceed with sensible defaults and mark unknown details as TBD.)"


# ---------- Tool Registry ----------
CORE_FUNCS = {
    # Files
    "list_dir": list_dir, "read_file": read_file, "search_files": search_files,
    "write_file": write_file, "move_file": move_file, "delete_file": delete_file,
    # Browser tools live in browser_tools.py
    **browser_tools.BROWSER_FUNCS,
}

ASK_SCHEMA = tool("ask_user", "Ask the user clarifying questions and wait. ONE call, short numbered list.",
                  {"question": "Numbered questions"}, ["question"])

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
    # File tools
    tool("list_dir", "List files in a workspace folder", ["path"], []),
    tool("read_file", "Read a small text file", ["path"], ["path"]),
    tool("search_files", "Find files by glob pattern", ["pattern"], ["pattern"]),
    tool("write_file", "Create or overwrite a text file", ["path", "content"], ["path", "content"]),
    tool("move_file", "Move or rename a file", ["src", "dst"], ["src", "dst"]),
    tool("delete_file", "Delete a file (moved to trash)", ["path"], ["path"]),
] + browser_tools.BROWSER_SCHEMAS

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
    "\nCRITICAL SEARCH RULE: Whenever the user asks to search, find websites, look up news, or check current info, "
    "you MUST call `search_web(query='...')` directly. "
    "NEVER call browser_open('https://google.com') and NEVER try to type or click into Google's search box manually. "
    "`search_web` automatically performs the search and returns results to you. "
    "Browser tools are named browser_* (browser_open, browser_click, browser_fill, browser_extract_text, ...). "
    "Keep replies short and never bypass confirmations."
)

OFFICE_PROMPT = (
    "\nDOCUMENT WORK: Pick Word (.docx) for reports/proposals, Excel (.xlsx) for budgets/data, "
    "PowerPoint (.pptx) for presentations. browser_export_pdf only when requested.\n\n"
    "CRITICAL POWERPOINT DESIGN RULES (NEVER MAKE GENERIC BULLET WALLS):\n"
    "A professional deck MUST use diverse graphic layouts across slides:\n"
    "1. TITLE SLIDE: Call ppt_start(title, subtitle, category).\n"
    "2. METRICS / NUMBERS: Call ppt_add_kpi_slide for key numbers, stats, and achievements (e.g. '$4.2M | Revenue | +25% YoY').\n"
    "3. PILLARS / KEY THEMES: Call ppt_add_cards_slide to present 3 or 4 strategic points in side-by-side modern cards.\n"
    "4. COMPARISONS / BEFORE & AFTER: Call ppt_add_split_slide for 2-column comparative analysis.\n"
    "5. DATA & TABLES: Call ppt_add_chart_slide for trends or ppt_add_table_slide for structured matrices.\n"
    "DO NOT put all content into plain bullets. Every deck must contain at least 2 visual card or KPI slides.\n"
)


# ---------- Token-saving helpers ----------
def compact(messages, tail=6, tool_tail=10, limit=1200):
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
    if len(history) <= keep + 1:
        return history
    tail = history[-keep:]
    for i, m in enumerate(tail):
        if m.get("role") == "user":
            return [history[0]] + tail[i:]
    return [history[0], history[-1]]


def _switch(model, why):
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
                    f"Every model hit its Groq limit. Try again in about {int(wait // 60) + 1} min.")
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
                raise RuntimeError("Request too large for Groq limit. Start a new chat.")
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
            browser_tools.close_browser()
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


# ---------- FastAPI Server ----------
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
    origin = ws.headers.get("origin")
    if origin and origin not in ALLOWED_ORIGINS:
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
    uvicorn.run(app, host="127.0.0.1", port=PORT)