import math
import os
import re
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches as PInches, Pt as PPt

from common import (
    approved,
    confirm,
    guard,
    parse_floats,
    parse_pipe_rows,
    safe,
    split_labels,
    tool,
)


# Executive Design System & Palette (McKinsey / Apple Keynote Style)

SLATE_900  = PRGB(0x0F, 0x17, 0x2A)   # #0F172A - Dominant dark
NAVY_CARD  = PRGB(0x1E, 0x29, 0x3B)   # #1E293B - Card fill / headers
BLUE_ACCENT= PRGB(0x25, 0x63, 0xEB)   # #2563EB - Royal Brand Blue
CYAN_ACCENT= PRGB(0x02, 0x84, 0xC7)   # #0284C7 - Accent Sky
LIGHT_BG   = PRGB(0xF8, 0xFA, 0xFC)   # #F8FAFC - Clean canvas
WHITE      = PRGB(0xFF, 0xFF, 0xFF)   # #FFFFFF - Crisp white cards
CARD_BG    = PRGB(0xF1, 0xF5, 0xF9)   # #F1F5F9 - Zebra row tint
BORDER     = PRGB(0xE2, 0xE8, 0xF0)   # #E2E8F0 - Ultra-subtle border
TEXT_DARK  = PRGB(0x0F, 0x17, 0x2A)   # #0F172A - High-contrast text
TEXT_MUTED = PRGB(0x64, 0x74, 0x8B)   # #64748B - Supporting text
SLATE_400  = PRGB(0x94, 0xA3, 0xB8)   # #94A3B8 - Metadata / Dark subtext

FONT_DISPLAY = "Segoe UI"
FONT_BODY    = "Segoe UI"

MARK = "pcagent-16x9"
SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5
MAX_TABLE_COLUMNS = 8
MAX_TABLE_ROWS = 12

# 
# Safe Persistence & Integrity
# 

def _validate_title(value, field="Title"):
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field} cannot be empty.")
    return value.strip()


def _save_presentation(prs, path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=".ppt_work_", suffix=".pptx", dir=str(path.parent))
    os.close(fd)
    try:
        prs.save(temp_name)
        check = Presentation(temp_name)
        if len(check.slides) == 0:
            raise ValueError("Generated presentation has no slides.")
        os.replace(temp_name, path)
    finally:
        if os.path.exists(temp_name):
            try:
                os.remove(temp_name)
            except OSError:
                pass


def _open_presentation(path):
    try:
        return Presentation(str(path))
    except Exception as exc:
        raise ValueError(f"Cannot open PowerPoint '{path}': {exc}") from exc


# 
# Graphic Primitives
# 

def _set_background(slide, color=WHITE):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def _add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _add_card_container(slide, left, top, width, height, bg_color=WHITE, border_color=BORDER):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = PPt(1)
    else:
        shape.line.fill.background()
    return shape


def _add_slide_header(slide, title, category="STRATEGY"):
    _set_background(slide, LIGHT_BG)

    # Category Pill
    tag_box = slide.shapes.add_textbox(PInches(0.8), PInches(0.4), PInches(6.0), PInches(0.3))
    p_tag = tag_box.text_frame.paragraphs[0]
    p_tag.margin_left = p_tag.margin_top = 0
    r_tag = p_tag.add_run()
    r_tag.text = category.upper()
    r_tag.font.name = FONT_DISPLAY
    r_tag.font.size = PPt(9)
    r_tag.font.bold = True
    r_tag.font.color.rgb = BLUE_ACCENT

    # Main Title
    title_box = slide.shapes.add_textbox(PInches(0.8), PInches(0.68), PInches(11.7), PInches(0.75))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = 0
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title
    r.font.name = FONT_DISPLAY
    r.font.size = PPt(24)
    r.font.bold = True
    r.font.color.rgb = SLATE_900

    # Header Divider Line
    _add_rect(slide, PInches(0.8), PInches(1.48), PInches(11.7), PInches(0.015), BORDER)


def _add_footer(slide, slide_number=None):
    _add_rect(slide, PInches(0.8), PInches(7.08), PInches(11.7), PInches(0.012), BORDER)
    if slide_number is not None:
        box = slide.shapes.add_textbox(PInches(11.7), PInches(7.12), PInches(0.8), PInches(0.25))
        p = box.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        r = p.add_run()
        r.text = str(slide_number)
        r.font.name = FONT_BODY
        r.font.size = PPt(8.5)
        r.font.color.rgb = TEXT_MUTED


# 
# Slide Builders
# 

def ppt_start(path: str, title: str, subtitle: str = "", category: str = "EXECUTIVE BRIEF") -> str:
    """Create a high-impact title slide with modern dark-mode typography."""
    title = _validate_title(title, "Presentation title")
    p = safe(path)
    if p.suffix.lower() != ".pptx":
        return "Error: path must end with .pptx"

    if not confirm(f"Create PowerPoint: {path}"):
        return "User denied."

    prs = Presentation()
    prs.slide_width = PInches(SLIDE_WIDTH)
    prs.slide_height = PInches(SLIDE_HEIGHT)
    prs.core_properties.comments = MARK

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_background(slide, SLATE_900)

    # Vertical left decorative accent
    _add_rect(slide, PInches(0), PInches(0), PInches(0.25), PInches(SLIDE_HEIGHT), BLUE_ACCENT)
    _add_rect(slide, PInches(0.25), PInches(0), PInches(0.08), PInches(SLIDE_HEIGHT), CYAN_ACCENT)

    # Category Pill
    pill = slide.shapes.add_textbox(PInches(1.2), PInches(2.0), PInches(8.0), PInches(0.4))
    r_pill = pill.text_frame.paragraphs[0].add_run()
    r_pill.text = category.upper()
    r_pill.font.name = FONT_DISPLAY
    r_pill.font.size = PPt(11)
    r_pill.font.bold = True
    r_pill.font.color.rgb = CYAN_ACCENT

    # Main Title
    t_box = slide.shapes.add_textbox(PInches(1.2), PInches(2.45), PInches(10.8), PInches(2.0))
    tf = t_box.text_frame
    tf.word_wrap = True
    r_title = tf.paragraphs[0].add_run()
    r_title.text = title
    r_title.font.name = FONT_DISPLAY
    r_title.font.size = PPt(40)
    r_title.font.bold = True
    r_title.font.color.rgb = WHITE

    # Subtitle
    if subtitle:
        sub_box = slide.shapes.add_textbox(PInches(1.2), PInches(4.5), PInches(10.5), PInches(1.2))
        sub_tf = sub_box.text_frame
        sub_tf.word_wrap = True
        r_sub = sub_tf.paragraphs[0].add_run()
        r_sub.text = subtitle
        r_sub.font.name = FONT_BODY
        r_sub.font.size = PPt(18)
        r_sub.font.color.rgb = SLATE_400

    _save_presentation(prs, p)
    approved.add(str(p))
    return f"Created modern presentation at {path}."


def ppt_add_slides(path: str, slides: str) -> str:
    """
    Intelligent Auto-Layout Engine:
    Detects bullet count and automatically formats into 2, 3, or 4 visual cards
    instead of a generic wall of text.
    """
    p = guard(path, ".pptx")
    prs = _open_presentation(p)

    # Parse markdown into slide blocks
    parsed = []
    current = None
    for line in slides.replace("\r\n", "\n").split("\n"):
        line_clean = line.strip()
        if not line_clean:
            continue
        if line.startswith("# "):
            current = {"title": line[2:].strip(), "items": [], "notes": []}
            parsed.append(current)
            continue
        if current is None:
            current = {"title": "Overview", "items": [], "notes": []}
            parsed.append(current)
        if line_clean.lower().startswith("notes:"):
            current["notes"].append(line_clean[6:].strip())
        else:
            bullet_clean = re.sub(r"^[-*\u2022]\s+", "", line_clean)
            if bullet_clean:
                current["items"].append(bullet_clean)

    for item in parsed:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _add_slide_header(slide, item["title"])
        items = item["items"]
        count = len(items)

        # -------------------------------------------------------------
        # AUTO-DESIGN 1: 2, 3, or 4 items -> Horizontal Visual Cards!
        # -------------------------------------------------------------
        if 2 <= count <= 4:
            total_w = 11.7
            gap = 0.3
            card_w = (total_w - (gap * (count - 1))) / count
            card_h = 4.85
            y = 1.85

            for idx, text in enumerate(items):
                x = 0.8 + idx * (card_w + gap)

                # Container card with top accent
                _add_card_container(slide, PInches(x), PInches(y), PInches(card_w), PInches(card_h), WHITE, BORDER)
                _add_rect(slide, PInches(x), PInches(y), PInches(card_w), PInches(0.08), BLUE_ACCENT)

                # Card Text Box
                tb = slide.shapes.add_textbox(PInches(x + 0.25), PInches(y + 0.35), PInches(card_w - 0.5), PInches(card_h - 0.6))
                tf = tb.text_frame
                tf.word_wrap = True

                # Badge / Step indicator (01, 02, 03)
                p_badge = tf.paragraphs[0]
                r_badge = p_badge.add_run()
                r_badge.text = f"0{idx + 1}"
                r_badge.font.name = FONT_DISPLAY
                r_badge.font.size = PPt(10)
                r_badge.font.bold = True
                r_badge.font.color.rgb = CYAN_ACCENT
                p_badge.space_after = PPt(8)

                # Parse lead-in title if formatted as "**Title**: body" or "Title: body"
                parts = re.split(r"[:\-–]\s*", text.replace("**", ""), 1)
                card_title = parts[0].strip() if len(parts) == 2 else f"Key Point {idx+1}"
                card_body = parts[1].strip() if len(parts) == 2 else text.replace("**", "")

                p_t = tf.add_paragraph()
                r_t = p_t.add_run()
                r_t.text = card_title
                r_t.font.name = FONT_DISPLAY
                r_t.font.size = PPt(18 if count <= 3 else 16)
                r_t.font.bold = True
                r_t.font.color.rgb = SLATE_900
                p_t.space_after = PPt(10)

                p_b = tf.add_paragraph()
                r_b = p_b.add_run()
                r_b.text = card_body
                r_b.font.name = FONT_BODY
                r_b.font.size = PPt(14 if count <= 3 else 13)
                r_b.font.color.rgb = TEXT_DARK

        # -------------------------------------------------------------
        # AUTO-DESIGN 2: 5 or more items -> 2-Column Grid of Cards!
        # -------------------------------------------------------------
        else:
            col_w = 5.7
            gap = 0.3
            y = 1.85
            half = math.ceil(count / 2)
            cols = [items[:half], items[half:]]

            for c_idx, col_items in enumerate(cols):
                cx = 0.8 + c_idx * (col_w + gap)
                _add_card_container(slide, PInches(cx), PInches(y), PInches(col_w), PInches(4.85), WHITE, BORDER)
                _add_rect(slide, PInches(cx), PInches(y), PInches(col_w), PInches(0.08), BLUE_ACCENT if c_idx == 0 else CYAN_ACCENT)

                tb = slide.shapes.add_textbox(PInches(cx + 0.3), PInches(y + 0.35), PInches(col_w - 0.6), PInches(4.3))
                tf = tb.text_frame
                tf.word_wrap = True

                for i, itm in enumerate(col_items):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.space_after = PPt(10)

                    # Bold lead-in
                    parts = itm.replace("**", "").split(":", 1)
                    if len(parts) == 2 and len(parts[0]) < 30:
                        r1 = p.add_run()
                        r1.text = "• " + parts[0].strip() + ": "
                        r1.font.bold = True
                        r1.font.color.rgb = SLATE_900
                        r1.font.name = FONT_DISPLAY
                        r1.font.size = PPt(14)

                        r2 = p.add_run()
                        r2.text = parts[1].strip()
                        r2.font.color.rgb = TEXT_MUTED
                        r2.font.name = FONT_BODY
                        r2.font.size = PPt(13)
                    else:
                        r = p.add_run()
                        r.text = "• " + itm.replace("**", "")
                        r.font.color.rgb = TEXT_DARK
                        r.font.name = FONT_BODY
                        r.font.size = PPt(13.5)

        if item["notes"]:
            slide.notes_slide.notes_text_frame.text = "\n".join(item["notes"])

        _add_footer(slide, len(prs.slides))

    _save_presentation(prs, p)
    return f"Added {len(parsed)} designed slide(s)."


def ppt_add_kpi_slide(path: str, title: str, kpis: str) -> str:
    """
    Add high-impact Hero KPI metric cards.
    kpis format: "Number | Label | Subtext"
    """
    p = guard(path, ".pptx")
    prs = _open_presentation(p)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_slide_header(slide, title, category="METRICS & PERFORMANCE")

    cards = [line.strip().split("|") for line in kpis.strip().split("\n") if line.strip()]
    count = min(len(cards), 4)
    if count == 0:
        return "Error: No KPI items provided."

    total_w = 11.7
    gap = 0.3
    card_w = (total_w - (gap * (count - 1))) / count
    y = 1.85

    for idx, card in enumerate(cards[:count]):
        metric = card[0].strip() if len(card) > 0 else ""
        label = card[1].strip() if len(card) > 1 else ""
        desc = card[2].strip() if len(card) > 2 else ""

        x = 0.8 + idx * (card_w + gap)
        _add_card_container(slide, PInches(x), PInches(y), PInches(card_w), PInches(4.85), WHITE, BORDER)
        _add_rect(slide, PInches(x), PInches(y), PInches(card_w), PInches(0.08), BLUE_ACCENT)

        box = slide.shapes.add_textbox(PInches(x + 0.25), PInches(y + 0.6), PInches(card_w - 0.5), PInches(3.8))
        tf = box.text_frame
        tf.word_wrap = True

        p_num = tf.paragraphs[0]
        r_num = p_num.add_run()
        r_num.text = metric
        r_num.font.name = FONT_DISPLAY
        r_num.font.size = PPt(38 if len(metric) <= 6 else 28)
        r_num.font.bold = True
        r_num.font.color.rgb = BLUE_ACCENT
        p_num.space_after = PPt(14)

        p_lbl = tf.add_paragraph()
        r_lbl = p_lbl.add_run()
        r_lbl.text = label
        r_lbl.font.name = FONT_DISPLAY
        r_lbl.font.size = PPt(17)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = SLATE_900
        p_lbl.space_after = PPt(8)

        if desc:
            p_desc = tf.add_paragraph()
            r_desc = p_desc.add_run()
            r_desc.text = desc
            r_desc.font.name = FONT_BODY
            r_desc.font.size = PPt(13.5)
            r_desc.font.color.rgb = TEXT_MUTED

    _add_footer(slide, len(prs.slides))
    _save_presentation(prs, p)
    return f"Added KPI cards slide with {count} metrics."


def ppt_add_table_slide(path: str, title: str, table: str) -> str:
    """Add a modern corporate table slide with clean padding and zebra rows."""
    p = guard(path, ".pptx")
    prs = _open_presentation(p)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_slide_header(slide, title, category="DATA SUMMARY")

    rows = parse_pipe_rows(table)
    if not rows or len(rows) < 2:
        return "Error: Table requires a header row and at least 1 data row."

    rows = [[str(cell).strip() for cell in r] for r in rows]
    cols = len(rows[0])
    row_count = min(len(rows), MAX_TABLE_ROWS)

    t_left = PInches(0.8)
    t_top = PInches(1.85)
    t_width = PInches(11.7)
    t_height = min(PInches(0.48 * row_count), PInches(4.9))

    shape = slide.shapes.add_table(row_count, cols, t_left, t_top, t_width, t_height)
    tbl = shape.table

    for i in range(cols):
        tbl.columns[i].width = int(t_width / cols)

    for i in range(row_count):
        row = rows[i]
        tbl.rows[i].height = int(t_height / row_count)
        is_header = (i == 0)

        for j in range(cols):
            val = row[j] if j < len(row) else ""
            cell = tbl.cell(i, j)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = PInches(0.15)

            cell.fill.solid()
            cell.fill.fore_color.rgb = SLATE_900 if is_header else (WHITE if i % 2 == 1 else CARD_BG)

            p = cell.text_frame.paragraphs[0]
            if not is_header and re.match(r"^[\$€£]?\s*[-+]?[0-9,.]+%?$", val):
                p.alignment = PP_ALIGN.RIGHT
            else:
                p.alignment = PP_ALIGN.LEFT

            for run in p.runs:
                run.font.name = FONT_DISPLAY if is_header else FONT_BODY
                run.font.size = PPt(13 if is_header else 12)
                run.font.bold = is_header
                run.font.color.rgb = WHITE if is_header else TEXT_DARK

    _add_footer(slide, len(prs.slides))
    _save_presentation(prs, p)
    return "Designed corporate table slide added."


def ppt_add_chart_slide(path: str, title: str, chart_type: str, labels: str, values: str, series_name: str = "Metric") -> str:
    """Add a native PowerPoint chart with executive palette styling."""
    p = guard(path, ".pptx")
    prs = _open_presentation(p)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_slide_header(slide, title, category="VISUAL ANALYSIS")

    chart_map = {
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }
    c_type = chart_map.get(chart_type.lower().strip(), XL_CHART_TYPE.COLUMN_CLUSTERED)

    labs = split_labels(labels)
    vals = parse_floats(values)
    if len(labs) != len(vals):
        return f"Error: {len(labs)} labels provided, but {len(vals)} values."

    data = CategoryChartData()
    data.categories = labs
    data.add_series(series_name, vals)

    c_shape = slide.shapes.add_chart(c_type, PInches(1.2), PInches(1.85), PInches(10.9), PInches(4.9), data)
    chart = c_shape.chart
    chart.has_title = False
    chart.has_legend = (chart_type.lower() == "pie")

    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.font.name = FONT_BODY
        chart.legend.font.size = PPt(10)

    try:
        if c_type in (XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED):
            chart.plots[0].gap_width = 75
            chart.series[0].format.fill.solid()
            chart.series[0].format.fill.fore_color.rgb = BLUE_ACCENT
    except Exception:
        pass

    _add_footer(slide, len(prs.slides))
    _save_presentation(prs, p)
    return f"Designed {chart_type} chart slide added."


# 
# Tool Registry
# 

REGISTRY = [
    (
        "ppt",
        tool(
            "ppt_start",
            "Create a new PowerPoint presentation title slide with modern dark-mode typography.",
            {
                "path": "File path (e.g. reports/Strategy.pptx)",
                "title": "Main presentation title",
                "subtitle": "Subtitle or summary line",
                "category": "Topic tag (e.g. EXECUTIVE BRIEF, STRATEGY REVIEW)",
            },
            ["path", "title"],
        ),
        ppt_start,
    ),
    (
        "ppt",
        tool(
            "ppt_add_slides",
            "Add content slides. Automatically formats bullets into 2, 3, or 4 modern horizontal cards.",
            {
                "path": "Existing .pptx path",
                "slides": "Slide markdown with '# Slide Title' followed by bullets (- **Lead-in**: description)",
            },
            ["path", "slides"],
        ),
        ppt_add_slides,
    ),
    (
        "ppt",
        tool(
            "ppt_add_kpi_slide",
            "Add high-impact KPI/Metric hero cards (e.g. $4.8M | Revenue | +24% YoY).",
            {
                "path": "Existing .pptx path",
                "title": "KPI Slide Title",
                "kpis": "Rows formatted as: 'Number | Label | Subtext'",
            },
            ["path", "title", "kpis"],
        ),
        ppt_add_kpi_slide,
    ),
    (
        "ppt",
        tool(
            "ppt_add_table_slide",
            "Add an executive corporate table slide with zebra rows.",
            {
                "path": "Existing .pptx path",
                "title": "Table Slide Title",
                "table": "Markdown pipe table: 'Col1 | Col2\\nVal1 | Val2'",
            },
            ["path", "title", "table"],
        ),
        ppt_add_table_slide,
    ),
    (
        "ppt",
        tool(
            "ppt_add_chart_slide",
            "Add a native PowerPoint chart (bar, column, line, or pie).",
            {
                "path": "Existing .pptx path",
                "title": "Chart Title",
                "chart_type": "bar, column, line, or pie",
                "labels": "Pipe-separated labels (Q1|Q2|Q3|Q4)",
                "values": "Pipe-separated numbers (120|145|180|210)",
                "series_name": "Series legend name",
            },
            ["path", "title", "chart_type", "labels", "values"],
        ),
        ppt_add_chart_slide,
    ),
]