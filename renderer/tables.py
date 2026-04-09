"""
renderer/tables.py — Styled table generation using python-pptx.

Tables are rendered with:
  - A bold, theme-colored header row (red background, white text)
  - Alternating light/white row backgrounds for readability
  - Consistent font sizes, padding, and column widths

Note: python-pptx's high-level cell fill API is limited, so cell
background colors are applied via direct XML manipulation.
"""

from __future__ import annotations

import logging

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn      # converts "a:tag" → "{namespace}tag" for lxml
from lxml import etree            # for direct XML manipulation of cell properties

import config
from renderer.utils import add_slide_title, add_slide_number, add_textbox

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Public entry point (called by engine.py)
# ---------------------------------------------------------------------------

def render_table_slide(slide, slide_data: dict, slide_num: int) -> None:
    """Render a complete table slide: title + table + optional caption + slide number.

    If the table fails to render (bad data, empty headers, etc.) a
    placeholder error message is shown instead.

    Args:
        slide: python-pptx slide object (already added to the presentation).
        slide_data: Single slide dict from the blueprint JSON.
        slide_num: Slide number to display in the bottom-right corner.
    """
    from pptx.enum.text import PP_ALIGN

    # --- Title ---
    title = slide_data.get("title", "")
    add_slide_title(slide, title)

    # Pull the table data dict and optional caption from the blueprint
    table_data = slide_data.get("table", {})
    caption = slide_data.get("caption", "")

    # Shrink table height if a caption needs to fit below it
    caption_height = Inches(0.4) if caption else 0
    table_height = config.CONTENT_HEIGHT - caption_height - config.ELEMENT_GAP

    # --- Table ---
    try:
        _add_styled_table(
            slide,
            table_data,
            left=config.MARGIN_LEFT,
            top=config.CONTENT_TOP,
            width=config.CONTENT_WIDTH,
            height=table_height,
        )
    except Exception as exc:
        # Show a visible error message if rendering fails
        logger.warning("Table render failed: %s", exc)
        add_textbox(
            slide, f"[Table could not be rendered: {exc}]",
            left=config.MARGIN_LEFT, top=config.CONTENT_TOP,
            width=config.CONTENT_WIDTH, height=table_height,
            font_size=config.BODY_FONT_SIZE, color=config.COLOR_TEXT_MUTED,
        )

    # --- Optional caption below the table ---
    if caption:
        add_textbox(
            slide, caption,
            left=config.MARGIN_LEFT,
            top=config.CONTENT_TOP + table_height + config.ELEMENT_GAP,
            width=config.CONTENT_WIDTH,
            height=Inches(0.35),
            font_size=config.CAPTION_FONT_SIZE,
            color=config.COLOR_TEXT_MUTED,
            align=PP_ALIGN.CENTER,
        )

    # --- Slide number ---
    add_slide_number(slide, slide_num)


# ---------------------------------------------------------------------------
# Internal table builder
# ---------------------------------------------------------------------------

def _add_styled_table(
    slide,
    table_data: dict,
    left: int,
    top: int,
    width: int,
    height: int,
) -> None:
    """Build a styled table shape and add it to the slide.

    Row 0 = styled header row (red background, white bold text, centered).
    Rows 1+ = data rows with alternating white/light-gray backgrounds.

    Args:
        slide: python-pptx slide object.
        table_data: Dict with "headers" (list[str]) and "rows" (list[list[str]]).
        left, top, width, height: Position and size in EMU.
    """
    headers = table_data.get("headers", [])
    rows = table_data.get("rows", [])

    if not headers:
        raise ValueError("Table has no headers.")   # can't build a table without column names

    num_cols = len(headers)
    num_rows = len(rows) + 1   # +1 accounts for the header row

    # Protect against tables with too many rows overflowing the slide
    if num_rows > 20:
        rows = rows[:19]      # keep only first 19 data rows
        num_rows = 20
        logger.debug("Table truncated to 20 rows to prevent overflow.")

    # python-pptx creates the table grid with all cells empty
    table_shape = slide.shapes.add_table(
        num_rows, num_cols,
        left, top, width, height,
    )
    table = table_shape.table

    # Distribute available width evenly across all columns
    col_width = width // num_cols
    for col in table.columns:
        col.width = col_width

    # --- Header row (row index 0) ---
    for col_idx, header_text in enumerate(headers):
        cell = table.cell(0, col_idx)    # row=0 is the header row
        cell.text = str(header_text)
        _style_cell(
            cell,
            font_size=Pt(13),
            bold=True,
            text_color=config.COLOR_TEXT_LIGHT,   # white text
            bg_color=config.COLOR_HEADER_BG,       # red background
            align=PP_ALIGN.CENTER,                 # centred in header
        )

    # --- Data rows (row indices 1 through num_rows-1) ---
    for row_idx, row_data in enumerate(rows):
        # Alternate background: odd rows get light gray, even rows stay white
        bg = config.COLOR_ALT_ROW if row_idx % 2 == 1 else config.COLOR_TEXT_LIGHT

        for col_idx in range(num_cols):
            cell = table.cell(row_idx + 1, col_idx)   # +1 because row 0 is the header

            # Use empty string for cells where the source row has fewer columns
            value = str(row_data[col_idx]) if col_idx < len(row_data) else ""
            cell.text = value

            _style_cell(
                cell,
                font_size=Pt(12),
                bold=False,
                text_color=config.COLOR_TEXT_DARK,
                bg_color=bg,
                align=PP_ALIGN.LEFT,   # body data is left-aligned
            )


def _style_cell(
    cell,
    font_size: Pt,
    bold: bool,
    text_color: RGBColor,
    bg_color: RGBColor,
    align: PP_ALIGN,
) -> None:
    """Apply font styling and background fill to a single table cell.

    python-pptx doesn't expose a direct cell.fill API that works reliably
    across PowerPoint versions, so we inject the fill via raw XML.

    Args:
        cell: A python-pptx _Cell object.
        font_size: Font size (Pt).
        bold: Whether text should be bold.
        text_color: RGB color for the text.
        bg_color: RGB color for the cell background fill.
        align: Paragraph text alignment.
    """
    # --- Background fill via XML ---
    # We add an <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill> element
    # inside the table cell properties element <a:tcPr>.
    try:
        tc = cell._tc                          # the underlying <a:tc> XML element
        tcPr = tc.get_or_add_tcPr()            # get or create <a:tcPr> (cell properties)

        # Build the solid fill XML structure
        solidFill = etree.SubElement(tcPr, qn("a:solidFill"))
        srgbClr = etree.SubElement(solidFill, qn("a:srgbClr"))

        # Convert RGBColor to uppercase 6-digit hex string (e.g. "E83F33")
        hex_color = "{:02X}{:02X}{:02X}".format(
            bg_color.red, bg_color.green, bg_color.blue
        )
        srgbClr.set("val", hex_color)
    except Exception as exc:
        logger.debug("Cell background fill failed: %s", exc)

    # --- Text formatting ---
    tf = cell.text_frame
    tf.word_wrap = True   # prevent text from overflowing into adjacent cells
    for para in tf.paragraphs:
        para.alignment = align
        for run in para.runs:
            run.font.size = font_size
            run.font.bold = bold
            run.font.color.rgb = text_color

    # --- Cell inner padding ---
    # Sets marL/R/T/B (margin left/right/top/bottom) on the tcPr element
    try:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        margin = Inches(0.05)   # small inner padding so text doesn't touch cell borders
        tcPr.set("marL", str(int(margin)))
        tcPr.set("marR", str(int(margin)))
        tcPr.set("marT", str(int(margin)))
        tcPr.set("marB", str(int(margin)))
    except Exception:
        pass   # padding is cosmetic — safe to skip on failure
