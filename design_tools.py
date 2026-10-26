"""Agent tools for the one-call restyle of EXISTING Word and PowerPoint files.

word_design_report -> word_design.py   (navy cover, contents, headings, tables, header/footer)
ppt_design_deck    -> ppt_design.py    (navy title slide and title bands, fonts, tables, charts, footer)

Both save a NEW file, never touch the original, and reopen the result to verify it.
(The Excel counterpart, excel_design_report, lives in excel_tools.py / excel_design.py.)
REGISTRY lists (group, schema, function) for main.py.
"""
from common import ROOT, approved, confirm, safe, tool
from ppt_design import design_deck
from word_design import design_document


def _prepare(path, output, ext):
    src = safe(path)
    if src.suffix.lower() != ext or not src.exists():
        raise ValueError(f"File not found: {path}")
    dst = safe(output) if output else src.with_name(f"{src.stem}_Designed{ext}")
    if dst.suffix.lower() != ext:
        raise ValueError(f"output must end with {ext}")
    if dst == src:
        raise ValueError("output must be a different file name; the original is kept untouched")
    if not confirm(f"Create designed copy: {dst.name} (original stays unchanged)"):
        return None, None
    return src, dst


def _summary(dst, res):
    rel = dst.relative_to(ROOT.resolve()).as_posix()
    status = "all checks passed" if res["ok"] else "SOME CHECKS FAILED"
    return f"Saved {rel} ({status}). " + "; ".join(res["checks"])


def _choice(value):
    v = (value or "auto").strip().lower()
    return v if v in ("auto", "yes", "no") else "auto"


def word_design_report(path: str, output: str = "", title: str = "", subtitle: str = "",
                       company: str = "", author: str = "", cover: str = "auto",
                       toc: str = "auto") -> str:
    src, dst = _prepare(path, output, ".docx")
    if src is None:
        return "User denied."
    res = design_document(src, dst, title, subtitle, company, author, _choice(cover), _choice(toc))
    approved.add(str(dst))
    return _summary(dst, res)


def ppt_design_deck(path: str, output: str = "", footer: str = "") -> str:
    src, dst = _prepare(path, output, ".pptx")
    if src is None:
        return "User denied."
    res = design_deck(src, dst, footer)
    approved.add(str(dst))
    return _summary(dst, res)


REGISTRY = [
    ("word", tool("word_design_report",
                  "ONE-CALL premium restyle of an existing Word document: navy cover, contents, styled headings and tables, header, 'Page X of Y'. Content is never changed. Saves a NEW file; original untouched; verifies by reopening. Use for any full redesign/restyle of a .docx.",
                  {"path": "Source .docx", "output": "New file, e.g. reports/Report_Designed.docx",
                   "title": "Cover title (blank = detect)", "subtitle": "Cover subtitle",
                   "company": "Company name", "author": "Author or team",
                   "cover": "auto, yes or no", "toc": "auto, yes or no"},
                  ["path", "output"]), word_design_report),
    ("ppt", tool("ppt_design_deck",
                 "ONE-CALL premium restyle of an existing PowerPoint: navy title slide and title bands, consistent fonts, styled tables and charts, footer and slide numbers. Content and notes are never changed. Saves a NEW file; original untouched; verifies by reopening. Use for any full redesign/restyle of a .pptx.",
                 {"path": "Source .pptx", "output": "New file, e.g. reports/Deck_Designed.pptx",
                  "footer": "Footer text (blank = deck title)"},
                 ["path", "output"]), ppt_design_deck),
]