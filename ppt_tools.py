import math
import os
import re
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.chart import XL_CHART_TYPE
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


# Design system

NAVY = PRGB(0x18, 0x2B, 0x49)
BLUE = PRGB(0x32, 0x6D, 0xC6)
CYAN = PRGB(0x36, 0xB9, 0xD6)
WHITE = PRGB(0xFF, 0xFF, 0xFF)
LIGHT_BG = PRGB(0xF4, 0xF7, 0xFB)
DARK_TEXT = PRGB(0x25, 0x32, 0x43)
MUTED_TEXT = PRGB(0x6B, 0x78, 0x8C)
BORDER = PRGB(0xDF, 0xE6, 0xEF)
PALE_BLUE = PRGB(0xE8, 0xF0, 0xFC)

MARK = "pcagent-16x9"
MAX_TABLE_COLUMNS = 8
MAX_TABLE_ROWS = 10

SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5


# Validation and safe save

def _validate_title(value, field="Title"):
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field} cannot be empty.")
    return value.strip()


def _save_presentation(prs, path):
    """
    Save to a temporary PPTX, verify it can be reopened,
    and atomically replace the destination.

    If saving fails, an existing destination is preserved.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    fd, temp_name = tempfile.mkstemp(
        prefix=".ppt_work_",
        suffix=".pptx",
        dir=str(path.parent),
    )
    os.close(fd)

    try:
        prs.save(temp_name)

        # Basic integrity validation.
        check = Presentation(temp_name)

        if len(check.slides) == 0:
            raise ValueError("Generated presentation has no slides.")

        # Validate that the PPTX archive can be read.
        with open(temp_name, "rb") as file:
            signature = file.read(4)

        if signature != b"PK\x03\x04":
            raise ValueError("Generated file is not a valid PPTX archive.")

        os.replace(temp_name, path)

    finally:
        if os.path.exists(temp_name):
            try:
                os.remove(temp_name)
            except OSError:
                pass


def _open_presentation(path):
    """Open an existing presentation with a useful error."""
    try:
        return Presentation(str(path))
    except Exception as exc:
        raise ValueError(
            f"Cannot open PowerPoint '{path}': {exc}"
        ) from exc


def _p16(prs):
    return prs.core_properties.comments == MARK


# Visual styling helpers

def _set_background(slide, color=WHITE):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def _add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left,
        top,
        width,
        height,
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _style_title(slide, font_size=28, color=NAVY):
    title = slide.shapes.title

    if title is None:
        return

    title.left = PInches(0.65)
    title.top = PInches(0.42)
    title.width = PInches(12.0)
    title.height = PInches(0.85)

    for paragraph in title.text_frame.paragraphs:
        paragraph.space_after = PPt(0)

        for run in paragraph.runs:
            run.font.name = "Aptos Display"
            run.font.size = PPt(font_size)
            run.font.bold = True
            run.font.color.rgb = color


def _add_footer(slide, slide_number=None):
    """Add a subtle footer and accent line."""
    _add_rect(
        slide,
        PInches(0.65),
        PInches(7.12),
        PInches(12.03),
        PInches(0.012),
        BORDER,
    )

    if slide_number is not None:
        box = slide.shapes.add_textbox(
            PInches(11.9),
            PInches(7.17),
            PInches(0.7),
            PInches(0.2),
        )
        paragraph = box.text_frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.RIGHT
        run = paragraph.add_run()
        run.text = str(slide_number)
        run.font.size = PPt(9)
        run.font.color.rgb = MUTED_TEXT


def _add_content_header(slide, title):
    _set_background(slide)
    _style_title(slide)
    _add_rect(
        slide,
        PInches(0.65),
        PInches(1.35),
        PInches(0.75),
        PInches(0.055),
        BLUE,
    )


def _fit(slide, prs, kind="content"):
    """Apply the design system only to presentations we create."""
    if not _p16(prs):
        return

    if kind == "title":
        title = slide.shapes.title

        if title is not None:
            title.left = PInches(0.9)
            title.top = PInches(2.0)
            title.width = PInches(11.5)
            title.height = PInches(1.5)

            for paragraph in title.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Aptos Display"
                    run.font.size = PPt(38)
                    run.font.bold = True
                    run.font.color.rgb = NAVY

        for placeholder in slide.placeholders:
            if placeholder.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
                placeholder.left = PInches(0.95)
                placeholder.top = PInches(3.8)
                placeholder.width = PInches(10.8)
                placeholder.height = PInches(1.0)

                for paragraph in placeholder.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Aptos"
                        run.font.size = PPt(19)
                        run.font.color.rgb = MUTED_TEXT

        return

    _add_content_header(slide, "")
    _add_footer(slide)


def _ppt_clear(prs):
    """Remove template slides before generating a new deck."""
    slide_ids = prs.slides._sldIdLst

    for slide_id in list(slide_ids):
        prs.part.drop_rel(slide_id.rId)
        slide_ids.remove(slide_id)


def _title_only_slide(prs, title):
    """Create a content slide with a professional title area."""
    if not prs.slide_layouts:
        raise ValueError("Presentation has no slide layouts.")

    idx = 5 if len(prs.slide_layouts) > 5 else len(prs.slide_layouts) - 1
    slide = prs.slides.add_slide(prs.slide_layouts[idx])

    if slide.shapes.title is not None:
        slide.shapes.title.text = title

    if _p16(prs):
        _add_content_header(slide, title)
        _add_footer(slide, len(prs.slides))

    return slide


def _body_placeholder(slide):
    """Find a suitable content placeholder without assuming index 1."""
    for placeholder in slide.placeholders:
        if placeholder.placeholder_format.type in (
            PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.OBJECT,
        ):
            return placeholder
    return None


# Create presentation

def ppt_start(
    path: str,
    title: str,
    subtitle: str = "",
    template: str = "",
) -> str:
    """Create a new designed PowerPoint."""

    title = _validate_title(title, "Presentation title")

    if not isinstance(subtitle, str):
        raise ValueError("Subtitle must be a string.")

    p = safe(path)

    if p.suffix.lower() != ".pptx":
        return "Error: path must end with .pptx"

    template_path = None

    if template:
        template_path = safe(template)

        if template_path.suffix.lower() != ".pptx":
            return "Error: template must be a .pptx file."

        if not template_path.is_file():
            return "Error: template file does not exist."

    if not confirm(f"Create PowerPoint: {path}"):
        return "User denied."

    if template_path is not None:
        prs = _open_presentation(template_path)
        _ppt_clear(prs)
    else:
        prs = Presentation()
        prs.slide_width = PInches(SLIDE_WIDTH)
        prs.slide_height = PInches(SLIDE_HEIGHT)
        prs.core_properties.comments = MARK

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    _set_background(slide, LIGHT_BG)

    if slide.shapes.title is not None:
        slide.shapes.title.text = title

    subtitle_placeholder = next(
        (
            ph for ph in slide.placeholders
            if ph.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE
        ),
        None,
    )

    if subtitle_placeholder is not None:
        subtitle_placeholder.text = subtitle

    if _p16(prs):
        _add_rect(
            slide,
            PInches(0),
            PInches(0),
            PInches(0.16),
            PInches(SLIDE_HEIGHT),
            BLUE,
        )

        _add_rect(
            slide,
            PInches(0.95),
            PInches(1.65),
            PInches(0.9),
            PInches(0.07),
            CYAN,
        )

    _fit(slide, prs, "title")

    try:
        _save_presentation(prs, p)
    except Exception as exc:
        return f"Error saving PowerPoint: {exc}"

    approved.add(str(p))

    return f"PowerPoint created at {path}. Add slides with ppt_add_slides."


# Parse slide content

def _parse_slides(md):
    if not isinstance(md, str) or not md.strip():
        raise ValueError("Slides input must be a non-empty string.")

    slides = []
    current = None

    for raw in md.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if not raw.strip():
            continue

        if raw.startswith("# "):
            title = raw[2:].strip()
            if not title:
                raise ValueError("Slide title cannot be empty.")

            current = {"title": title, "items": [], "notes": []}
            slides.append(current)
            continue

        if current is None:
            raise ValueError(
                "Start slide content with '# Slide title'."
            )

        stripped = raw.strip()

        if stripped.lower().startswith("notes:"):
            note = stripped[6:].strip()
            if note:
                current["notes"].append(note)
            continue

        level = 1 if raw.startswith(("  ", "\t")) else 0

        item = re.sub(r"^[-*\u2022]\s+", "", stripped).strip()
        item = item.replace("**", "")

        if item:
            current["items"].append((level, item))

    if not slides:
        raise ValueError("No slides found.")

    return slides


# Add bullet slides

def ppt_add_slides(path: str, slides: str) -> str:
    """Add formatted bullet slides to an existing presentation."""
    p = guard(path, ".pptx")
    prs = _open_presentation(p)
    parsed = _parse_slides(slides)

    layout = (
        prs.slide_layouts[1]
        if len(prs.slide_layouts) > 1
        else prs.slide_layouts[0]
    )

    has_body = any(
        ph.placeholder_format.type in (
            PP_PLACEHOLDER.BODY,
            PP_PLACEHOLDER.OBJECT,
        )
        for ph in layout.placeholders
    )

    if any(sd["items"] for sd in parsed) and not has_body:
        raise ValueError(
            "Selected slide layout has no body placeholder."
        )

    for data in parsed:
        slide = prs.slides.add_slide(layout)

        if slide.shapes.title is not None:
            slide.shapes.title.text = data["title"]

        body = _body_placeholder(slide)

        if body is not None:
            body.left = PInches(0.8)
            body.top = PInches(1.7)
            body.width = PInches(11.7)
            body.height = PInches(5.1)

            tf = body.text_frame
            tf.clear()
            tf.word_wrap = True
            tf.margin_left = PInches(0.12)
            tf.margin_right = PInches(0.12)

            for index, (level, text) in enumerate(data["items"]):
                paragraph = (
                    tf.paragraphs[0]
                    if index == 0
                    else tf.add_paragraph()
                )

                paragraph.text = text
                paragraph.level = level
                paragraph.space_after = PPt(12 if level == 0 else 6)

                for run in paragraph.runs:
                    run.font.name = "Aptos"
                    run.font.size = PPt(20 if level == 0 else 16)
                    run.font.color.rgb = DARK_TEXT

        if data["notes"]:
            slide.notes_slide.notes_text_frame.text = "\n".join(
                data["notes"]
            )

        if _p16(prs):
            _add_content_header(slide, data["title"])
            _add_footer(slide, len(prs.slides))

    try:
        _save_presentation(prs, p)
    except Exception as exc:
        return f"Error saving presentation: {exc}"

    return f"Added {len(parsed)} slide(s)."


# Add table slide

def ppt_add_table_slide(path: str, title: str, table: str) -> str:
    """Add a professionally formatted table."""

    p = guard(path, ".pptx")
    title = _validate_title(title, "Table title")
    prs = _open_presentation(p)

    rows = parse_pipe_rows(table)

    if not rows or len(rows) < 2:
        raise ValueError(
            "Table needs a header and at least one data row."
        )

    rows = [[str(cell).strip() for cell in row] for row in rows]
    columns = len(rows[0])

    if columns == 0:
        raise ValueError("Table must have at least one column.")

    if columns > MAX_TABLE_COLUMNS:
        raise ValueError(
            f"Maximum supported columns: {MAX_TABLE_COLUMNS}."
        )

    if any(len(row) != columns for row in rows):
        raise ValueError(
            "All table rows must contain the same number of columns."
        )

    if len(rows) > MAX_TABLE_ROWS:
        raise ValueError(
            f"Maximum supported rows: {MAX_TABLE_ROWS}. "
            "Split the table across slides."
        )

    slide = _title_only_slide(prs, title)

    left = PInches(0.65)
    top = PInches(1.75)
    width = prs.slide_width - PInches(1.3)
    height = min(
        PInches(0.55) * len(rows),
        prs.slide_height - PInches(2.65),
    )

    shape = slide.shapes.add_table(
        len(rows), columns, left, top, width, height
    )

    tbl = shape.table

    # Equal column widths initially; adjust later if needed.
    for col in tbl.columns:
        col.width = int(width / columns)

    row_height = int(height / len(rows))

    for i, row in enumerate(rows):
        tbl.rows[i].height = row_height

        for j, value in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = value
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = PInches(0.10)
            cell.margin_right = PInches(0.10)
            cell.margin_top = PInches(0.04)
            cell.margin_bottom = PInches(0.04)

            if i == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = NAVY
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = (
                    WHITE if i % 2 else LIGHT_BG
                )

            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT

                for run in paragraph.runs:
                    run.font.name = "Aptos"
                    run.font.size = PPt(
                        13 if columns <= 4 else 10
                    )
                    run.font.bold = i == 0
                    run.font.color.rgb = (
                        WHITE if i == 0 else DARK_TEXT
                    )

    try:
        _save_presentation(prs, p)
    except Exception as exc:
        return f"Error saving table slide: {exc}"

    return "Designed table slide added."


# Add chart slide

def ppt_add_chart_slide(
    path: str,
    title: str,
    chart_type: str,
    labels: str,
    values: str,
    series_name: str = "",
) -> str:
    """Add a validated, styled native PowerPoint chart."""

    p = guard(path, ".pptx")
    title = _validate_title(title, "Chart title")
    prs = _open_presentation(p)

    chart_name = (
        chart_type.strip().lower()
        if isinstance(chart_type, str)
        else ""
    )

    chart_types = {
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }

    if chart_name not in chart_types:
        raise ValueError(
            "Chart type must be bar, column, line, or pie."
        )

    labs = split_labels(labels)
    vals = parse_floats(values)

    if not labs or not vals:
        raise ValueError("Chart needs labels and numeric values.")

    if len(labs) != len(vals):
        raise ValueError(
            f"Found {len(labs)} labels but {len(vals)} values."
        )

    if any(not str(label).strip() for label in labs):
        raise ValueError("Chart labels cannot be empty.")

    if not all(math.isfinite(value) for value in vals):
        raise ValueError("Chart values must be finite numbers.")

    if chart_name == "pie":
        if any(value < 0 for value in vals):
            raise ValueError("Pie chart values cannot be negative.")
        if sum(vals) <= 0:
            raise ValueError("Pie chart total must be greater than zero.")

    series_name = (
        series_name.strip()
        if isinstance(series_name, str) and series_name.strip()
        else "Series"
    )

    data = CategoryChartData()
    data.categories = [str(label).strip() for label in labs]
    data.add_series(series_name, vals)

    slide = _title_only_slide(prs, title)

    chart_shape = slide.shapes.add_chart(
        chart_types[chart_name],
        PInches(0.8),
        PInches(1.65),
        prs.slide_width - PInches(1.6),
        prs.slide_height - PInches(2.35),
        data,
    )

    chart = chart_shape.chart
    chart.has_title = False
    chart.has_legend = chart_name == "pie"

    if chart_name != "pie":
        chart.has_legend = False
        chart.value_axis.has_major_gridlines = True
        chart.category_axis.tick_labels.font.name = "Aptos"
        chart.category_axis.tick_labels.font.size = PPt(11)
        chart.value_axis.tick_labels.font.name = "Aptos"
        chart.value_axis.tick_labels.font.size = PPt(10)

    if chart_name in ("bar", "column"):
        plot = chart.plots[0]
        plot.gap_width = 65

    # Apply the primary accent color to the series where supported.
    if chart_name != "pie":
        try:
            series = chart.series[0]
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = BLUE
        except (AttributeError, IndexError, ValueError):
            pass

    try:
        _save_presentation(prs, p)
    except Exception as exc:
        return f"Error saving chart slide: {exc}"

    return f"Designed {chart_name} chart added ({len(labs)} points)."

# Registry - compatible with main.py

REGISTRY = [
    (
        "ppt",
        tool(
            "ppt_start",
            "Create a designed PowerPoint title slide, optionally from a template.",
            {
                "path": "e.g. reports/Deck.pptx",
                "title": "Presentation title",
                "subtitle": "Optional subtitle",
                "template": "Optional .pptx template",
            },
            ["path", "title"],
        ),
        ppt_start,
    ),
    (
        "ppt",
        tool(
            "ppt_add_slides",
            "Add slides using '# Slide title', '- bullet', indented sub-bullets, and optional 'Notes: ...'.",
            {
                "path": "Existing .pptx path",
                "slides": "Slide content",
            },
            ["path", "slides"],
        ),
        ppt_add_slides,
    ),
    (
        "ppt",
        tool(
            "ppt_add_table_slide",
            "Add a styled table slide. Separate cells using |.",
            {
                "path": "Existing .pptx path",
                "title": "Table slide title",
                "table": "Header and data rows separated by |",
            },
            ["path", "title", "table"],
        ),
        ppt_add_table_slide,
    ),
    (
        "ppt",
        tool(
            "ppt_add_chart_slide",
            "Add a styled native bar, column, line, or pie chart.",
            {
                "path": "Existing .pptx path",
                "title": "Chart title",
                "chart_type": "bar, column, line, or pie",
                "labels": "A|B|C",
                "values": "1|2|3",
                "series_name": "Optional series name",
            },
            ["path", "title", "chart_type", "labels", "values"],
        ),
        ppt_add_chart_slide,
    ),
]
