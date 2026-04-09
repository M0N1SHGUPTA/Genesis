"""
renderer/layouts.py — Layout rendering functions for content slides.

Each function receives a slide object + slide_data dict and draws all
shapes at precise grid-aligned positions using constants from config.py.

How it works:
  render_content_slide() reads the "layout" field from the blueprint and
  dispatches to the correct private function (_two_column, _three_cards, etc.).
  If a layout fails, it silently falls back to _single_focus.

Layouts implemented:
  two_column, three_cards, key_stats, timeline,
  process_flow, comparison, icon_list, single_focus
"""

from __future__ import annotations

import logging

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

import config
from renderer.utils import (
    add_textbox,
    add_bullet_textbox,
    add_slide_number,
    add_slide_title,
    style_shape,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Short aliases for frequently used config values
# Keeps the coordinate arithmetic below compact and readable.
# ---------------------------------------------------------------------------
ML = config.MARGIN_LEFT     # left edge of the usable content area
CT = config.CONTENT_TOP     # top edge of the body area (below title)
CW = config.CONTENT_WIDTH   # total usable width
CH = config.CONTENT_HEIGHT  # total usable height (below title, above bottom margin)
GAP = config.ELEMENT_GAP    # vertical gap between stacked elements
PAD = config.INNER_PADDING  # padding inside cards/boxes
CGAP = config.CARD_GAP      # horizontal gap between side-by-side cards


# ---------------------------------------------------------------------------
# Public dispatcher
# ---------------------------------------------------------------------------

def render_content_slide(slide, slide_data: dict, slide_num: int) -> None:
    """Entry point for all non-special content slides.

    Reads the "layout" key from slide_data, adds the title,
    then calls the matching private layout function.

    Args:
        slide: python-pptx slide object (already added to the presentation).
        slide_data: Single slide dict from the blueprint JSON.
        slide_num: Slide number shown in the bottom-right corner.
    """
    layout = slide_data.get("layout", "single_focus")   # default if LLM omits the field
    title = slide_data.get("title", "")

    # Title is added first so all layout functions can assume it's already there
    add_slide_title(slide, title)

    # Map layout name strings to their render functions
    dispatch = {
        "two_column":   _two_column,
        "three_cards":  _three_cards,
        "key_stats":    _key_stats,
        "timeline":     _timeline,
        "process_flow": _process_flow,
        "comparison":   _comparison,
        "icon_list":    _icon_list,
        "single_focus": _single_focus,
    }

    fn = dispatch.get(layout, _single_focus)   # unknown layout → single_focus fallback
    try:
        fn(slide, slide_data)
    except Exception as exc:
        # If the specific layout crashes, degrade gracefully to the simplest layout
        logger.warning("Layout '%s' render error: %s — falling back to single_focus", layout, exc)
        try:
            _single_focus(slide, slide_data)
        except Exception:
            pass   # absolute last resort: leave the slide blank rather than crash

    # Always add a slide number last so it sits on top of any other shapes
    add_slide_number(slide, slide_num)


# ---------------------------------------------------------------------------
# Layout: two_column
# Splits the content area into two equal side-by-side cards.
# Each card has a heading + bullet list.
# ---------------------------------------------------------------------------

def _two_column(slide, data: dict) -> None:
    """Render a two-column layout.

    Blueprint keys expected:
        left.heading  (str)
        left.points   (list[str])
        right.heading (str)
        right.points  (list[str])
    """
    # Each column gets half the content width minus a gap in the middle
    col_width = (CW - CGAP) / 2
    col_height = CH - GAP   # leave a small gap at the bottom

    for i, side in enumerate(("left", "right")):
        col = data.get(side, {})
        if not isinstance(col, dict):
            continue   # skip if blueprint data is malformed

        # Column left edge: first column starts at margin, second is offset by one column + gap
        left = ML + i * (col_width + CGAP)

        # Cream-colored card background with red border
        card = slide.shapes.add_shape(1, left, CT, col_width, col_height)
        style_shape(card, fill_color=config.COLOR_CARD_BG, line_color=config.COLOR_CARD_BORDER)

        # Red bold heading inside the card (with inner padding from card edge)
        heading = col.get("heading", "")
        if heading:
            add_textbox(
                slide, heading,
                left=left + PAD, top=CT + PAD,
                width=col_width - 2 * PAD, height=Inches(0.4),
                font_size=config.CARD_HEADING_SIZE, bold=True,
                color=config.COLOR_PRIMARY,
            )

        # Bullet points below the heading
        points = col.get("points", [])
        if points:
            add_bullet_textbox(
                slide, points,
                left=left + PAD,
                top=CT + PAD + Inches(0.5),   # start below the heading
                width=col_width - 2 * PAD,
                height=col_height - PAD - Inches(0.6),
                font_size=config.BODY_FONT_SIZE,
                color=config.COLOR_TEXT_DARK,
            )


# ---------------------------------------------------------------------------
# Layout: three_cards
# Three equal cards in a row, each with a number badge, heading, and bullets.
# ---------------------------------------------------------------------------

def _three_cards(slide, data: dict) -> None:
    """Render a three-cards layout.

    Blueprint key expected:
        cards  (list of {number, heading, points})
    """
    cards = data.get("cards", [])[:3]   # cap at 3 — layout only has room for 3
    if not cards:
        _single_focus(slide, data)   # graceful fallback if no card data
        return

    # Each card gets a third of the width minus the two inter-card gaps
    card_width = (CW - 2 * CGAP) / 3
    card_height = CH - GAP

    for i, card_data in enumerate(cards):
        # Each card's left edge is offset by its index × (card width + gap)
        left = ML + i * (card_width + CGAP)

        # Cream background card with red border
        card = slide.shapes.add_shape(1, left, CT, card_width, card_height)
        style_shape(card, fill_color=config.COLOR_CARD_BG, line_color=config.COLOR_CARD_BORDER)

        # Large red number badge in the top-left of the card (e.g. "01", "02", "03")
        number = card_data.get("number", str(i + 1).zfill(2))
        add_textbox(
            slide, number,
            left=left + PAD, top=CT + PAD,
            width=card_width - 2 * PAD, height=Inches(0.5),
            font_size=Pt(28), bold=True,
            color=config.COLOR_PRIMARY,
        )

        # Card heading in dark bold text, below the number
        heading = card_data.get("heading", "")
        if heading:
            add_textbox(
                slide, heading,
                left=left + PAD, top=CT + PAD + Inches(0.55),
                width=card_width - 2 * PAD, height=Inches(0.5),
                font_size=config.CARD_HEADING_SIZE, bold=True,
                color=config.COLOR_TEXT_DARK,
            )

        # Bullet points below the heading, smaller font to fit in the card
        points = card_data.get("points", [])
        if points:
            add_bullet_textbox(
                slide, points,
                left=left + PAD,
                top=CT + PAD + Inches(1.15),   # below heading
                width=card_width - 2 * PAD,
                height=card_height - PAD - Inches(1.25),
                font_size=Pt(12),   # slightly smaller than body to fit more text
                color=config.COLOR_TEXT_DARK,
            )


# ---------------------------------------------------------------------------
# Layout: key_stats
# 2–4 large numbers each with a label beneath — for displaying statistics.
# ---------------------------------------------------------------------------

def _key_stats(slide, data: dict) -> None:
    """Render 2–4 big stat numbers with descriptive labels.

    Blueprint key expected:
        stats  (list of {value, label})
    """
    stats = data.get("stats", [])[:4]   # max 4 stats fit horizontally
    if not stats:
        _single_focus(slide, data)
        return

    n = len(stats)
    # Distribute stats evenly across the content width
    stat_width = (CW - (n - 1) * CGAP) / n
    stat_top = CT + Inches(0.5)       # push down a bit from title for visual breathing room
    stat_height = CH - Inches(1.0)    # leave space at top and bottom

    for i, stat in enumerate(stats):
        left = ML + i * (stat_width + CGAP)

        # Cream card box for each stat
        box = slide.shapes.add_shape(1, left, stat_top, stat_width, stat_height)
        style_shape(box, fill_color=config.COLOR_CARD_BG, line_color=config.COLOR_CARD_BORDER)

        # The big number — 44pt bold red, centred in the card
        value = str(stat.get("value", ""))
        add_textbox(
            slide, value,
            left=left + PAD, top=stat_top + Inches(0.4),
            width=stat_width - 2 * PAD, height=Inches(1.2),
            font_size=config.STAT_NUMBER_SIZE, bold=True,
            color=config.COLOR_PRIMARY,
            align=PP_ALIGN.CENTER,
        )

        # The descriptive label below the number — smaller, muted gray
        label = str(stat.get("label", ""))
        add_textbox(
            slide, label,
            left=left + PAD, top=stat_top + Inches(1.7),
            width=stat_width - 2 * PAD, height=Inches(0.5),
            font_size=config.STAT_LABEL_SIZE,
            color=config.COLOR_TEXT_MUTED,
            align=PP_ALIGN.CENTER,
        )


# ---------------------------------------------------------------------------
# Layout: timeline
# Horizontal line with numbered circle nodes and text below each node.
# ---------------------------------------------------------------------------

def _timeline(slide, data: dict) -> None:
    """Render a horizontal timeline with numbered steps.

    Blueprint key expected:
        steps  (list of {number, heading, description})
    """
    steps = data.get("steps", [])[:5]   # max 5 steps fit on a horizontal timeline
    if not steps:
        _single_focus(slide, data)
        return

    n = len(steps)
    step_width = (CW - (n - 1) * CGAP) / n   # width allocated to each step
    circle_r = Inches(0.3)                    # radius of the circle node
    line_y = CT + Inches(1.0)                 # y-coordinate of the horizontal line
    circle_top = line_y - circle_r            # top of each circle (centred on the line)
    circle_size = circle_r * 2                # diameter

    # Draw the full-width horizontal connector line behind all circles
    line = slide.shapes.add_shape(
        1,
        ML, line_y - Inches(0.02),   # slight upward offset to visually centre on line_y
        CW, Inches(0.04),            # thin horizontal bar
    )
    style_shape(line, fill_color=config.COLOR_PRIMARY, line_color=None)

    for i, step in enumerate(steps):
        # cx = horizontal centre of this step's circle
        cx = ML + i * (step_width + CGAP) + step_width / 2

        # Red filled circle node sitting on the line
        circle = slide.shapes.add_shape(
            9,   # shape type 9 = oval
            cx - circle_r, circle_top,
            circle_size, circle_size,
        )
        style_shape(circle, fill_color=config.COLOR_PRIMARY, line_color=None)

        # Step number displayed as white text inside the circle
        number = step.get("number", str(i + 1).zfill(2))
        add_textbox(
            slide, number,
            left=cx - circle_r, top=circle_top,
            width=circle_size, height=circle_size,
            font_size=Pt(11), bold=True,
            color=config.COLOR_TEXT_LIGHT,
            align=PP_ALIGN.CENTER,
        )

        # Step heading in bold dark text, positioned below the circle
        heading = step.get("heading", "")
        add_textbox(
            slide, heading,
            left=ML + i * (step_width + CGAP),
            top=line_y + circle_r + Inches(0.1),   # just below the circle's bottom edge
            width=step_width, height=Inches(0.4),
            font_size=Pt(13), bold=True,
            color=config.COLOR_TEXT_DARK,
            align=PP_ALIGN.CENTER,
        )

        # Description text below the heading, muted gray
        desc = step.get("description", "")
        add_textbox(
            slide, desc,
            left=ML + i * (step_width + CGAP),
            top=line_y + circle_r + Inches(0.55),
            width=step_width, height=Inches(1.5),
            font_size=Pt(11),
            color=config.COLOR_TEXT_MUTED,
            align=PP_ALIGN.CENTER,
        )


# ---------------------------------------------------------------------------
# Layout: process_flow
# Boxes connected by arrows left-to-right — good for workflows / sequences.
# ---------------------------------------------------------------------------

def _process_flow(slide, data: dict) -> None:
    """Render a left-to-right process flow with step boxes and arrow connectors.

    Blueprint key expected:
        steps  (list of {number, heading, description})
    """
    steps = data.get("steps", [])[:5]
    if not steps:
        _single_focus(slide, data)
        return

    n = len(steps)
    arrow_w = Inches(0.3)                    # width of each arrow connector
    total_arrow = (n - 1) * arrow_w          # total space consumed by all arrows
    box_width = (CW - total_arrow) / n       # each step box gets an equal share
    box_height = Inches(2.2)
    # Centre the row of boxes vertically in the content area
    box_top = CT + (CH - box_height) / 2

    for i, step in enumerate(steps):
        # Each box starts immediately after the previous box + arrow
        left = ML + i * (box_width + arrow_w)

        # Step box with cream fill and red border
        box = slide.shapes.add_shape(1, left, box_top, box_width, box_height)
        style_shape(box, fill_color=config.COLOR_CARD_BG, line_color=config.COLOR_PRIMARY)

        # Large red step number at the top of the box
        number = step.get("number", str(i + 1).zfill(2))
        add_textbox(
            slide, number,
            left=left + PAD, top=box_top + PAD,
            width=box_width - 2 * PAD, height=Inches(0.4),
            font_size=Pt(20), bold=True, color=config.COLOR_PRIMARY,
            align=PP_ALIGN.CENTER,
        )

        # Step heading below the number
        heading = step.get("heading", "")
        add_textbox(
            slide, heading,
            left=left + PAD, top=box_top + Inches(0.55),
            width=box_width - 2 * PAD, height=Inches(0.45),
            font_size=Pt(13), bold=True, color=config.COLOR_TEXT_DARK,
            align=PP_ALIGN.CENTER,
        )

        # Short description at the bottom of the box
        desc = step.get("description", "")
        add_textbox(
            slide, desc,
            left=left + PAD, top=box_top + Inches(1.1),
            width=box_width - 2 * PAD, height=Inches(1.0),
            font_size=Pt(11), color=config.COLOR_TEXT_MUTED,
            align=PP_ALIGN.CENTER,
        )

        # Draw a thin red arrow bar between this box and the next
        # (not drawn after the last box)
        if i < n - 1:
            arrow_left = left + box_width   # arrow starts at the right edge of the box
            arrow = slide.shapes.add_shape(
                1,
                arrow_left,
                box_top + box_height / 2 - Inches(0.04),   # vertically centred on the boxes
                arrow_w, Inches(0.08),
            )
            style_shape(arrow, fill_color=config.COLOR_PRIMARY, line_color=None)


# ---------------------------------------------------------------------------
# Layout: comparison
# Two solid-colored columns (red vs dark) for direct side-by-side comparison.
# ---------------------------------------------------------------------------

def _comparison(slide, data: dict) -> None:
    """Render a high-contrast side-by-side comparison with two full-color columns.

    Blueprint keys expected:
        left.heading, left.points, right.heading, right.points
    """
    col_width = (CW - CGAP) / 2
    col_height = CH - GAP

    # Left column = primary red, Right column = near-black
    colors = [config.COLOR_PRIMARY, config.COLOR_TEXT_DARK]

    for i, side in enumerate(("left", "right")):
        col = data.get(side, {})
        if not isinstance(col, dict):
            continue

        left = ML + i * (col_width + CGAP)
        bg = colors[i]   # alternate background color per column

        # Solid filled column (no border — the color contrast is the visual separator)
        box = slide.shapes.add_shape(1, left, CT, col_width, col_height)
        style_shape(box, fill_color=bg, line_color=None)

        # Both columns use white text since both backgrounds are dark
        text_color = config.COLOR_TEXT_LIGHT

        heading = col.get("heading", "")
        if heading:
            add_textbox(
                slide, heading,
                left=left + PAD, top=CT + PAD,
                width=col_width - 2 * PAD, height=Inches(0.5),
                font_size=config.CARD_HEADING_SIZE, bold=True,
                color=text_color, align=PP_ALIGN.CENTER,
            )

        points = col.get("points", [])
        if points:
            add_bullet_textbox(
                slide, points,
                left=left + PAD, top=CT + PAD + Inches(0.6),
                width=col_width - 2 * PAD,
                height=col_height - PAD - Inches(0.7),
                font_size=config.BODY_FONT_SIZE, color=text_color,
            )


# ---------------------------------------------------------------------------
# Layout: icon_list
# Numbered circle + heading + description for each row — like a feature list.
# ---------------------------------------------------------------------------

def _icon_list(slide, data: dict) -> None:
    """Render a numbered icon-circle list (3–4 rows).

    Blueprint key expected:
        items  (list of {number, heading, description})
    """
    items = data.get("items", [])[:4]
    if not items:
        _single_focus(slide, data)
        return

    # Calculate row height dynamically based on how many items there are
    row_height = min(CH / len(items) - CGAP, Inches(1.3))
    circle_size = Inches(0.55)                        # diameter of the numbered circle
    text_left = ML + circle_size + Inches(0.25)       # text starts to the right of the circles
    text_width = CW - circle_size - Inches(0.25)      # text width = remaining space

    for i, item in enumerate(items):
        row_top = CT + i * (row_height + CGAP / 2)
        # Vertically centre the circle within the row
        circle_top = row_top + (row_height - circle_size) / 2

        # Red filled circle on the left
        circle = slide.shapes.add_shape(9, ML, circle_top, circle_size, circle_size)
        style_shape(circle, fill_color=config.COLOR_PRIMARY, line_color=None)

        # White number inside the circle
        number = item.get("number", str(i + 1).zfill(2))
        add_textbox(
            slide, number,
            left=ML, top=circle_top,
            width=circle_size, height=circle_size,
            font_size=Pt(14), bold=True, color=config.COLOR_TEXT_LIGHT,
            align=PP_ALIGN.CENTER,
        )

        # Bold heading to the right of the circle
        heading = item.get("heading", "")
        add_textbox(
            slide, heading,
            left=text_left, top=row_top,
            width=text_width, height=Inches(0.35),
            font_size=config.CARD_HEADING_SIZE, bold=True,
            color=config.COLOR_TEXT_DARK,
        )

        # Muted description text below the heading
        desc = item.get("description", "")
        add_textbox(
            slide, desc,
            left=text_left, top=row_top + Inches(0.38),
            width=text_width, height=row_height - Inches(0.4),
            font_size=config.BODY_FONT_SIZE,
            color=config.COLOR_TEXT_MUTED,
        )

        # Thin red separator line below each row (skip the last row)
        if i < len(items) - 1:
            line = slide.shapes.add_shape(
                1,
                ML, row_top + row_height + CGAP / 4,
                CW, Inches(0.01),   # 0.01" thin line
            )
            style_shape(line, fill_color=config.COLOR_CARD_BORDER, line_color=None)


# ---------------------------------------------------------------------------
# Layout: single_focus
# One large key statement + supporting bullet points below it.
# Used as the default fallback for unknown/missing layouts.
# ---------------------------------------------------------------------------

def _single_focus(slide, data: dict) -> None:
    """Render a single-focus slide: one large statement + supporting bullet list.

    Blueprint keys expected:
        focus   (str)  — the main key message displayed large
        points  (list[str]) — supporting bullet points below
    """
    focus = data.get("focus", "")
    points = data.get("points", [])

    if focus:
        # Large red focus statement at the top of the content area
        add_textbox(
            slide, focus,
            left=ML, top=CT,
            width=CW, height=Inches(1.2),
            font_size=Pt(24), bold=True,
            color=config.COLOR_PRIMARY,
            align=PP_ALIGN.LEFT,
        )

        # Thin red accent line below the focus statement as a visual separator
        line = slide.shapes.add_shape(
            1, ML, CT + Inches(1.25), Inches(1.5), Inches(0.05)
        )
        style_shape(line, fill_color=config.COLOR_PRIMARY, line_color=None)

        # Bullets start below the focus statement + accent line
        bullet_top = CT + Inches(1.4)
    else:
        # No focus text — bullets start at the top of the content area
        bullet_top = CT

    if points:
        add_bullet_textbox(
            slide, points,
            left=ML, top=bullet_top,
            width=CW,
            height=CH - (bullet_top - CT) - GAP,   # fill remaining vertical space
            font_size=config.BODY_FONT_SIZE,
            color=config.COLOR_TEXT_DARK,
        )
