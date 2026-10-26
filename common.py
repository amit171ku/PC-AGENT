"""Shared helpers for all document tools (Word, Excel, PowerPoint, PDF)."""
import os
import re
from pathlib import Path

# The ONLY folder the agent can touch. Override with the AGENT_WORKSPACE environment variable.
ROOT = Path(os.getenv("AGENT_WORKSPACE") or (Path.home() / "AgentWorkspace"))
TRASH = ROOT / ".trash"
LOG = ROOT / "agent_log.txt"
for _d in (ROOT, TRASH, ROOT / "templates", ROOT / "reports"):
    _d.mkdir(parents=True, exist_ok=True)

approved = set()          # files the user already approved editing in this session
_handler = {"fn": None}   # set by main.py: asks the user in the browser


def set_confirm_handler(fn):
    _handler["fn"] = fn


def confirm(action: str) -> bool:
    """Ask the user to allow an action. Denies when no handler is registered."""
    fn = _handler["fn"]
    return bool(fn(action)) if fn else False


def safe(path: str) -> Path:
    """Resolve a path and make sure it stays inside the workspace."""
    p = (ROOT / path).resolve()
    if ROOT.resolve() not in p.parents and p != ROOT.resolve():
        raise PermissionError(f"Path outside workspace: {path}")
    return p


def guard(path: str, ext: str) -> Path:
    """Return an existing file of the right type, asking the user once before editing it."""
    p = safe(path)
    if p.suffix.lower() != ext or not p.exists():
        raise ValueError(f"File not found: {path}. Create it first with its start tool.")
    if str(p) not in approved:
        if not confirm(f"Edit file: {path}"):
            raise PermissionError("User denied.")
        approved.add(str(p))
    return p


def parse_pipe_rows(text: str):
    """Parse pipe-separated text (also markdown tables) into a list of rows."""
    rows = []
    for line in text.replace("\r", "").split("\n"):
        s = line.strip()
        if not s:
            continue
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|"):
            s = s[:-1]
        cells = [c.strip() for c in s.split("|")]
        if cells and all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
            continue
        rows.append(cells)
    return rows


def parse_floats(text: str):
    return [float(v.strip().replace(",", "")) for v in text.split("|") if v.strip()]


def split_labels(text: str):
    return [v.strip() for v in text.split("|") if v.strip()]


def tool(name, desc, props, required):
    """Build an OpenAI-style tool schema. props: list of names or {name: description}."""
    if isinstance(props, list):
        props = {k: "" for k in props}
    properties = {k: ({"type": "string", "description": d} if d else {"type": "string"})
                  for k, d in props.items()}
    return {"type": "function", "function": {
        "name": name, "description": desc,
        "parameters": {"type": "object", "properties": properties, "required": required}}}