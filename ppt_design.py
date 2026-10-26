"""
Premium navy-and-white design for an EXISTING PowerPoint (.pptx) deck.

Run directly:
    python ppt_design.py "C:\\path\\Deck.pptx"
    python ppt_design.py <source.pptx> <output.pptx> --footer "ACME | Q1 Review"

The source file is never modified. The restyled copy keeps every slide, text, table, chart,
picture and speaker note; only the look changes: navy title slide, navy title band on content
slides, consistent fonts and bullet sizes, styled tables and charts, footer and slide numbers.
The saved file is reopened and checked.
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

NAVY, BLUE, LINE, ZEBRA = "0B2545", "3A7CA5", "D5DDE8", "F3F6FB"
INK, MUTED, PALE, WHITE = "1B2A41", "5B6B82", "A9C4E4", "FFFFFF"
FONT = "Calibri"
PALETTE = ["0B2545", "3A7CA5", "7FB0E5", "A9C4E4", "5B6B82", "D5DDE8"]
MARK = "pcagent-16x9"          # decks created by ppt_tools carry this marker
PREFIX = "Design "             # names of the shapes this script adds (so reruns can replace them)

_SKIP_TYPES = (PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE)


def _rgb(h):
    return RGBColor.from_string(h)


# ---------------------------------------------------------------- shape helpers
def _rect(slide, left, top, width, height, color, name):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(color)
    shp.line.fill.background()
    shp.shadow.inherit = False
    shp.name = PREFIX + name
    return shp


def _to_back(slide, shp):
    tree = slide.shapes._spTree
    tree.remove(shp._element)
    tree.insert(2, shp._element)


def _remove(shape):
    shape._element.getparent().remove(shape._element)


def _is_decor(shape, marked, W, H):
    """Thin bars/lines and slide-number boxes that ppt_tools adds; only stripped on its own decks."""
    if not marked:
        return False
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and not shape.has_text_frame:
            return min(shape.width, shape.height) < Inches(0.3)
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and shape.has_text_frame \
                and not shape.text_frame.text.strip():
            return min(shape.width, shape.height) < Inches(0.3)
        if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX and shape.has_text_frame:
            t = shape.text_frame.text.strip()
            return t.isdigit() and shape.top is not None and shape.top > H - Inches(0.7)
    except Exception:
        return False
    return False


def _style_runs(tf, size_by_level=None, size=None, color=INK, bold=None):
    for p in tf.paragraphs:
        sz = size_by_level.get(p.level, size_by_level.get(max(size_by_level))) if size_by_level else size
        for r in p.runs:
            r.font.name = FONT
            if sz:
                r.font.size = Pt(sz)
            if color:
                r.font.color.rgb = _rgb(color)
            if bold is not None:
                r.font.bold = bold


def _fit_text(tf):
    tf.word_wrap = True
    try:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:
        pass


# ---------------------------------------------------------------- tables and charts
def _cell_border(cell, color, w=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for old in tcPr.findall(qn(tag)):
            tcPr.remove(old)
    for i, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
        ln = etree.Element(qn(tag), w=str(w), cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
        etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
        tcPr.insert(i, ln)


def _style_table(shape, sc):
    tbl = shape.table
    ncols = len(tbl.columns)
    size = round((18 if ncols <= 4 else 15 if ncols <= 6 else 12) * sc)
    for ri, row in enumerate(tbl.rows):
        for cell in row.cells:
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(NAVY if ri == 0 else (ZEBRA if ri % 2 == 0 else WHITE))
            _cell_border(cell, LINE)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.1)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.name = FONT
                    r.font.size = Pt(size)
                    r.font.bold = True if ri == 0 else r.font.bold
                    r.font.color.rgb = _rgb(WHITE if ri == 0 else INK)


def _style_chart(chart, sc):
    try:
        chart.font.name = FONT
        chart.font.size = Pt(round(14 * sc))
        chart.font.color.rgb = _rgb(INK)
        pie = chart.chart_type in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.PIE_EXPLODED,
                                   XL_CHART_TYPE.DOUGHNUT, XL_CHART_TYPE.THREE_D_PIE)
        line = "LINE" in str(chart.chart_type)
        for si, ser in enumerate(chart.plots[0].series):
            if pie:
                for pi in range(len(list(ser.values))):
                    pt = ser.points[pi]
                    pt.format.fill.solid()
                    pt.format.fill.fore_color.rgb = _rgb(PALETTE[pi % len(PALETTE)])
            elif line:
                ser.format.line.color.rgb = _rgb(PALETTE[si % 3])
                ser.format.line.width = Pt(2.5)
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = _rgb(PALETTE[si % 3])
    except Exception:
        pass


# ---------------------------------------------------------------- footer
def _footer(slide, W, H, text, number, sc, add_text, add_number):
    _to_back(slide, _rect(slide, Inches(0.6), H - Inches(0.6), W - Inches(1.2), Inches(0.012),
                          LINE, "Footer Line"))
    if add_text and text:
        tb = slide.shapes.add_textbox(Inches(0.6), H - Inches(0.52), int(W * 0.6), Inches(0.3))
        tb.name = PREFIX + "Footer"
        tb.text_frame.word_wrap = True
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = text
        r.font.name, r.font.size = FONT, Pt(round(11 * sc))
        r.font.color.rgb = _rgb(MUTED)
    if add_number:
        tb = slide.shapes.add_textbox(W - Inches(1.6), H - Inches(0.52), Inches(1.0), Inches(0.3))
        tb.name = PREFIX + "Number"
        tb.text_frame.word_wrap = True
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        fld = etree.SubElement(p._p, qn("a:fld"))
        fld.set("id", "{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}")
        fld.set("type", "slidenum")
        rpr = etree.SubElement(fld, qn("a:rPr"))
        rpr.set("lang", "en-US")
        rpr.set("sz", str(int(round(11 * sc) * 100)))
        sf = etree.SubElement(rpr, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", MUTED)
        etree.SubElement(rpr, qn("a:latin")).set("typeface", FONT)
        etree.SubElement(fld, qn("a:t")).text = str(number)


# ---------------------------------------------------------------- slide designs
def _is_title_slide(slide, idx):
    name = (slide.slide_layout.name or "").lower()
    return idx == 1 or "title slide" in name or "section header" in name


def _design_title_slide(slide, W, H, sc):
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = _rgb(NAVY)
    title = slide.shapes.title
    sub = next((ph for ph in slide.placeholders
                if ph.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE
                or (ph.placeholder_format.idx == 1 and ph != title)), None)
    t_top = int(H * 0.28)
    t_h = Inches(1.7)
    if title is not None:
        title.left, title.top, title.width, title.height = Inches(0.9), t_top, W - Inches(1.8), t_h
        tf = title.text_frame
        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
        _fit_text(tf)
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.LEFT
        _style_runs(tf, size=round(46 * sc), color=WHITE, bold=True)
    bar_top = t_top + t_h + Inches(0.15)
    _rect(slide, Inches(0.9), bar_top, Inches(1.3), Inches(0.07), BLUE, "Title Accent")
    if sub is not None and sub.has_text_frame:
        sub.left, sub.top, sub.width, sub.height = Inches(0.9), bar_top + Inches(0.3), W - Inches(1.8), Inches(1.1)
        tf = sub.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        _fit_text(tf)
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.LEFT
        _style_runs(tf, size=round(24 * sc), color=PALE)


def _design_content_slide(slide, idx, W, H, sc, footer_text):
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = _rgb(WHITE)
    band_h = Inches(1.0) if sc >= 0.95 else Inches(0.9)
    title = slide.shapes.title
    existing = {ph.placeholder_format.type for ph in slide.placeholders}

    added = [_rect(slide, 0, 0, W, band_h, NAVY, "Band"),
             _rect(slide, 0, band_h, W, Inches(0.05), BLUE, "Accent")]
    for shp in added:
        _to_back(slide, shp)

    if title is not None:
        title.left, title.top, title.width, title.height = Inches(0.6), Inches(0.08), W - Inches(1.2), band_h - Inches(0.16)
        tf = title.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        _fit_text(tf)
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.LEFT
        _style_runs(tf, size=round(32 * sc), color=WHITE, bold=True)

    content_top = band_h + Inches(0.3)
    footer_top = H - Inches(0.7)
    bodies = [ph for ph in slide.placeholders
              if ph != title and ph.placeholder_format.type not in _SKIP_TYPES and ph.has_text_frame]
    levels = {0: round(26 * sc), 1: round(22 * sc), 2: round(19 * sc), 3: round(17 * sc)}
    for ph in bodies:
        tf = ph.text_frame
        _fit_text(tf)
        for p in tf.paragraphs:
            p.space_after = Pt(8)
        _style_runs(tf, size_by_level=levels, color=INK)
    single = bodies[0] if len(bodies) == 1 else None

    for sh in list(slide.shapes):
        if sh == title or sh.name.startswith(PREFIX) or sh.top is None or sh.height is None:
            continue
        if sh.is_placeholder and sh.placeholder_format.type in _SKIP_TYPES:
            continue
        if sh == single:
            sh.left, sh.top = Inches(0.7), content_top
            sh.width, sh.height = W - Inches(1.4), footer_top - content_top - Inches(0.1)
            continue
        if sh.top < content_top - Inches(0.05):
            sh.top = content_top
        if sh.top + sh.height > footer_top:
            room = max(footer_top - sh.top, Inches(0.5))
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                ratio = room / sh.height
                sh.width, sh.height = int(sh.width * ratio), int(room)
            else:
                sh.height = int(room)
        if getattr(sh, "has_table", False) and sh.has_table:
            _style_table(sh, sc)
        elif getattr(sh, "has_chart", False) and sh.has_chart:
            _style_chart(sh.chart, sc)
        elif sh.has_text_frame and not sh.is_placeholder and sh.text_frame.text.strip():
            _style_runs(sh.text_frame, color=None)

    _footer(slide, W, H, footer_text, idx, sc,
            add_text=PP_PLACEHOLDER.FOOTER not in existing,
            add_number=PP_PLACEHOLDER.SLIDE_NUMBER not in existing)


# ---------------------------------------------------------------- verification helpers
def _snapshot(prs, marked):
    W, H = prs.slide_width, prs.slide_height
    out = []
    for slide in prs.slides:
        texts, pics, charts, tables = [], 0, 0, 0
        for sh in slide.shapes:
            if sh.name.startswith(PREFIX) or _is_decor(sh, marked, W, H):
                continue
            if sh.has_text_frame and sh.text_frame.text.strip():
                texts.append(sh.text_frame.text.strip())
            if getattr(sh, "has_table", False) and sh.has_table:
                tables += 1
                for row in sh.table.rows:
                    for c in row.cells:
                        if c.text.strip():
                            texts.append(c.text.strip())
            if getattr(sh, "has_chart", False) and sh.has_chart:
                charts += 1
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pics += 1
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
        out.append({"texts": Counter(texts), "pics": pics, "charts": charts,
                    "tables": tables, "notes": notes})
    return out


def _soffice():
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if exe:
        return exe
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(c):
            return c
    return None


def _pdf_pages(path):
    try:
        from pypdf import PdfReader
        return len(PdfReader(str(path)).pages)
    except Exception:
        return len(re.findall(rb"/Type\s*/Page[^s]", Path(path).read_bytes()))


# ---------------------------------------------------------------- main entry
def design_deck(src, dst, footer=""):
    src, dst = Path(src), Path(dst)
    if src.resolve() == dst.resolve():
        raise ValueError("Output must be a different file. The original is kept untouched.")
    prs = Presentation(str(src))
    if len(prs.slides) == 0:
        raise ValueError("The presentation has no slides.")
    marked = prs.core_properties.comments == MARK
    W, H = prs.slide_width, prs.slide_height
    sc = min(1.0, W / Inches(13.333))
    before = _snapshot(prs, marked)

    first = prs.slides[0]
    footer_text = footer.strip() or (
        first.shapes.title.text.strip() if first.shapes.title is not None and first.shapes.title.text.strip()
        else src.stem.replace("_", " "))

    for idx, slide in enumerate(prs.slides, start=1):
        for sh in list(slide.shapes):                      # replace earlier design shapes / old decor
            if sh.name.startswith(PREFIX) or _is_decor(sh, marked, W, H):
                _remove(sh)
        if _is_title_slide(slide, idx):
            _design_title_slide(slide, W, H, sc)
        else:
            _design_content_slide(slide, idx, W, H, sc, footer_text)

    dst.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(dst))
    checks, ok = verify(src, dst, before, marked, len(prs.slides), (W, H))
    return {"output": str(dst), "slides": len(prs.slides), "checks": checks, "ok": ok}


def verify(src, dst, before, marked, n_slides, size):
    checks, ok = [], True
    prs = Presentation(str(dst))
    after = _snapshot(prs, marked)

    if len(prs.slides) != n_slides or (prs.slide_width, prs.slide_height) != size:
        ok = False
        checks.append("SLIDE COUNT OR SIZE CHANGED")
    problems = []
    for i, (b, a) in enumerate(zip(before, after), start=1):
        if (b["texts"] - a["texts"]) or b["notes"] != a["notes"] or \
                (b["pics"], b["charts"], b["tables"]) != (a["pics"], a["charts"], a["tables"]):
            problems.append(i)
    if not problems:
        total = sum(sum(b["texts"].values()) for b in before)
        pics = sum(b["pics"] for b in before)
        charts = sum(b["charts"] for b in before)
        tables = sum(b["tables"] for b in before)
        checks.append(f"content intact ({n_slides} slides, {total} text blocks, {tables} tables, "
                      f"{charts} charts, {pics} pictures, speaker notes compared)")
    else:
        ok = False
        checks.append(f"CONTENT MISMATCH on slide(s) {problems}")

    numbered = sum(1 for s in prs.slides
                   if any(sh.name == PREFIX + "Number" or
                          (sh.is_placeholder and sh.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER)
                          for sh in s.shapes))
    banded = sum(1 for s in prs.slides if any(sh.name == PREFIX + "Band" for sh in s.shapes))
    checks.append(f"{banded} content slides with navy title band, {numbered} with slide numbers, "
                  f"{n_slides - banded} title slide(s) in navy")

    exe = _soffice()
    if exe:
        try:
            with tempfile.TemporaryDirectory() as td:
                subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", td, str(dst)],
                               check=True, timeout=180, capture_output=True)
                pages = _pdf_pages(Path(td) / (Path(dst).stem + ".pdf"))
                if pages == n_slides:
                    checks.append(f"opens and renders ({pages} slides)")
                else:
                    ok = False
                    checks.append(f"RENDER MISMATCH: {pages} pages for {n_slides} slides")
        except Exception as e:
            checks.append(f"render check skipped ({str(e)[:50]})")
    else:
        checks.append("render check skipped (LibreOffice not found); open the file in PowerPoint to confirm")
    return checks, ok


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Apply a premium navy design to a PowerPoint deck.")
    ap.add_argument("source")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--footer", default="")
    a = ap.parse_args()
    source = Path(a.source)
    if not source.exists():
        sys.exit(f"File not found: {source}")
    target = Path(a.output) if a.output else source.with_name(source.stem + "_Designed.pptx")
    res = design_deck(source, target, a.footer)
    print("Saved:", res["output"])
    for line in res["checks"]:
        print(" -", line)
    print("ALL CHECKS PASSED" if res["ok"] else "SOME CHECKS FAILED")