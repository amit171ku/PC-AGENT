"""Run from your project folder:  python ui_patch.py   (or: python ui_patch.py path/to/static)

Patches static/app.js, static/style.css and static/index.html.
A .bak backup of each changed file is saved next to it. Safe to run twice.
"""
import shutil
import sys
from pathlib import Path

static = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("static")


def sub(s, old, new):
    if s.count(old) != 1:
        raise SystemExit("Could not find this text exactly once, so nothing was changed:\n" + old[:90])
    return s.replace(old, new)


def edit(name, fn):
    p = static / name
    if not p.exists():
        print(f"! {p} not found")
        return
    old = p.read_text(encoding="utf-8")
    new = fn(old)
    if new == old:
        print(f"= {name}: already up to date")
        return
    shutil.copy(p, p.with_name(p.name + ".bak"))
    p.write_text(new, encoding="utf-8")
    print(f"+ {name}: updated (backup: {p.name}.bak)")


# ---------------------------------------------------------------- app.js
CONFIRM_BODY = r'''// Turns "Key: value" lines into a tidy label / value grid. Falls back to plain text.
function confirmBody(text, box) {
  const body = document.createElement("div");
  const lines = String(text).split("\n").filter((l) => l.trim());
  const rows = lines.map((l) => /^([A-Za-z ]{2,14}):\s+(.+)$/.exec(l));
  if (lines.length > 1 && rows.every(Boolean)) {
    body.className = "c-rows";
    rows.forEach((m) => {
      const k = document.createElement("span");
      k.className = "c-key";
      k.textContent = m[1];
      const v = document.createElement("span");
      v.className = "c-val";
      v.textContent = m[2];
      body.append(k, v);
    });
    if (/may SUBMIT or CHANGE/i.test(text)) box.classList.add("risky");
  } else {
    body.className = "c-body";
    body.textContent = text;
  }
  return body;
}

'''

STATUS_OLD = '''    search_web: "Searching the web...",
    open_url: "Opening the page...",
    get_page_text: "Reading the page...",
    click: "Clicking on the page...",
    type_text: "Typing on the page...",
'''
STATUS_NEW = '''    search_web: "Searching the web...",
    browser_open: "Opening the page...",
    browser_navigate: "Navigating the browser...",
    browser_snapshot: "Looking at the page...",
    browser_click_ref: "Clicking on the page...",
    browser_fill_ref: "Typing on the page...",
    browser_click: "Clicking on the page...",
    browser_fill: "Typing on the page...",
    browser_select: "Choosing an option...",
    browser_check: "Ticking a checkbox...",
    browser_press_key: "Pressing a key...",
    browser_scroll: "Scrolling the page...",
    browser_wait: "Waiting for the page...",
    browser_extract_text: "Reading the page...",
    browser_extract_links: "Collecting links...",
    browser_extract_table: "Reading a table...",
    browser_screenshot: "Taking a screenshot...",
    browser_export_pdf: "Saving the page as PDF...",
'''


def patch_app(s):
    if "function confirmBody" in s:
        return s
    s = sub(s, "function addConfirm(id, text) {", CONFIRM_BODY + "function addConfirm(id, text) {")
    s = sub(s,
            '  const body = document.createElement("div");\n  body.className = "c-body";\n  body.textContent = text;\n',
            "  const body = confirmBody(text, box);\n")
    s = sub(s, STATUS_OLD, STATUS_NEW)
    return s


# ---------------------------------------------------------------- style.css
CSS_MARK = "/* Approval card: label / value rows */"
CSS_ADD = '''

''' + CSS_MARK + '''
.confirm .c-rows {
  display: grid; grid-template-columns: max-content 1fr; gap: 6px 14px;
  background: rgba(255, 255, 255, 0.55); border-radius: 8px;
  padding: 10px 12px; font-size: 13.5px; line-height: 1.45;
}
:root[data-theme="dark"] .confirm .c-rows { background: rgba(0, 0, 0, 0.25); }
.confirm .c-key { color: var(--muted); font-weight: 600; }
.confirm .c-val { color: var(--text); word-break: break-word; white-space: pre-wrap; }
.confirm.risky { border-color: var(--err); box-shadow: 0 0 0 2px color-mix(in srgb, var(--err) 18%, transparent); }
.confirm.risky .c-head { color: var(--err); }
.confirm.risky.done { border-color: var(--warn-border); box-shadow: none; }
.confirm.risky.done .c-head { color: var(--muted); }
'''


def patch_css(s):
    if CSS_MARK in s:
        return s
    return s.rstrip("\n") + CSS_ADD


# ---------------------------------------------------------------- index.html
def patch_html(s):
    if "Open google.com in a new browser tab" not in s and "Open a website for search" not in s:
        return s
    s = sub(s, 'data-prompt="Open google.com in a new browser tab"',
            'data-prompt="Open https://news.ycombinator.com and list the top 5 stories with their links"')
    s = sub(s, "</span> Open a website for search", "</span> Open a website")
    s = sub(s, 'data-prompt="Search the web your topic and save a short summary in tips.txt"',
            'data-prompt="Search the web for the latest news on AI agents, read the top 2 pages, and save a short summary with source links in tips.txt"')
    return s


edit("app.js", patch_app)
edit("style.css", patch_css)
edit("index.html", patch_html)
print("Done. Hard-refresh the browser page (Ctrl+F5) to see the changes.")