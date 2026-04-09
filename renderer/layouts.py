"""
renderer/layouts.py — Layout rendering functions for content slides.

Each function receives a slide object + slide_data dict and draws all
shapes at precise grid-aligned positions using constants from config.py.

How it works:
  render_content_slide() reads the "layout" field from the blueprint and
  dispatches to the correct private function (_two_column, _three_cards, etc.).
  Agenda slides get their own dedicated layout.
  If a layout fails, it silently falls back to _single_focus.

Layouts implemented:
  two_column, three_cards, key_stats, timeline,
  process_flow, comparison, icon_list, single_focus, agenda
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
# ---------------------------------------------------------------------------
ML = config.MARGIN_LEFT     # left edge of the usable content area
CT = config.CONTENT_TOP     # top edge of the body area (below title)
CW = config.CONTENT_WIDTH   # total usable width
CH = config.CONTENT_HEIGHT  # total usable height (below title, above bottom margin)
GAP = config.ELEMENT_GAP    # vertical gap between stacked elements
PAD = config.INNER_PADDING  # padding inside cards/boxes
CGAP = config.CARD_GAP      # horizontal gap between side-by-side cards


def _truncate(text: str, max_words: int) -> str:
    """Trim text to max_words words, appending '…' if trimmed."""
    words = text.split()
    if len(words) <= max_words:
        return text
    return " ".join(words[:max_words]) + "…"


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
    layout = slide_data.get("layout", "single_focus")
    title = slide_data.get("title", "")
    slide_type = slide_data.get("type", "content")

    # Title is added first so all layout functions can assume it's already there
    add_slide_title(slide, title)

    # Agenda gets its own dedicated renderer regardless of layout field
    if slide_type == "agenda":
        try:
            _agenda_layout(slide, slide_data)
        except Exception as exc:
            logger.warning("Agenda layout error: %s", exc)
            _single_focus(slide, slide_data)
        add_slide_number(slide, slide_num)
        return

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

    fn = dispatch.get(layout, _single_focus)
    try:
        fn(slide, slide_data)
    except Exception as exc:
        logger.warning("Layout '%s' render error: %s — falling back to single_focus", layout, exc)
        try:
            _single_focus(slide, slide_data)
        except Exception:
            pass

    add_slide_number(slide, slide_num)


# ---------------------------------------------------------------------------
# Layout: agenda
# Two-column list of all section headings, numbered.
# ---------------------------------------------------------------------------

def _agenda_layout(slide, data: dict) -> None:
    """Render an agenda/TOC slide.

    Splits items into two columns when there are more than 4 items.
    Uses numbered entries instead of bullet dots.

    Blueprint key expected:
        points  (list[str]) — list of section headings
    """
    points = data.get("points", [])
    if not points:
        return

    # Cap to avoid overflow — 12 items max (6 per column)
    points = points[:12]

    ITEM_FONT = Pt(15)
    NUM_FONT = Pt(18)
    ITEM_HEIGHT = Inches(0.65)
    NUM_WIDTH = Inches(0.55)

    if len(points) <= 4:
        # Single column, large numbered items
        for i, point in enumerate(points):
            row_top = CT + i * (ITEM_HEIGHT + Inches(0.1))
            # Red number badge
            num_box = slide.shapes.add_shape(9, ML, row_top + Inches(0.05),
                                             Inches(0.45), Inches(0.45))
            style_shape(num_box, fill_color=config.COLOR_PRIMARY, line_color=None)
            add_textbox(
                slide, str(i + 1).zfill(2),
                left=ML, top=row_top + Inches(0.05),
                width=Inches(0.45), height=Inches(0.45),
                font_size=NUM_FONT, bold=True,
                color=config.COLOR_TEXT_LIGHT,
                align=PP_ALIGN.CENTER,
            )
            # Item text
            add_textbox(
                slide, _truncate(point, 12),
                left=ML + Inches(0.6), top=row_top,
                width=CW - Inches(0.6), height=ITEM_HEIGHT,
                font_size=ITEM_FONT, color=config.COLOR_TEXT_DARK,
            )
    else:
        # Two columns — split items in half
        mid = (len(points) + 1) // 2
        left_pts = points[:mid]
        right_pts = points[mid:]

        col_width = (CW - CGAP) / 2
        col_height = CH - GAP

        for col_i, pts in enumerate([left_pts, right_pts]):
            col_left = ML + col_i * (col_width + CGAP)

            # Subtle card background
            card = slide.shapes.add_shape(1, col_left, CT, col_width, col_height)
            style_shape(card, fill_color=config.COLOR_CARD_BG, line_color=config.COLOR_CARD_BORDER)

            for row_i, point in enumerate(pts[:6]):
                global_num = col_i * mid + row_i + 1
                row_top = CT + PAD + row_i * (ITEM_HEIGHT + Inches(0.05))

                # Red number
                add_textbox(
                    slide, str(global_num).zfill(2),
                    left=col_left + PAD, top=row_top,
                    width=NUM_WIDTH, height=ITEM_HEIGHT,
                    font_size=Pt(14), bold=True,
                    color=config.COLOR_PRIMARY,
                    align=PP_ALIGN.CENTER,
                )
                # Heading text
                add_textbox(
                    slide, _truncate(point, 10),
                    left=col_left + PAD + NUM_WIDTH + Inches(0.1),
                    top=row_top,
                    width=col_width - PAD - NUM_WIDTH - Inches(0.2),
                    height=ITEM_HEIGHT,
                    font_size=Pt(13),
                    color=config.COLOR_TEXT_DARK,
                )


# ---------------------------------------------------------------------------
# Layout: two_column
# Splits the content area into two equal side-by-side cards.
# ---------------------------------------------------------------------------

def _two_column(slide, data: dict) -> None:
    """Render a two-column layout.

    Blueprint keys expected:
        left.heading  (str)
        left.points   (list[str])
        right.heading (str)
        right.points  (list[str])
    """
    col_width = (CW - CGAP) / 2
    col_height = CH - GAP

    for i, side in enumerate(("left", "right")):
        col = data.get(side, {})
        if not isinstance(col, dict):
            col = {}

        left = ML + i * (col_width + CGAP)

        # Cream card background with red border
        card = slide.shapes.add_shape(1, left, CT, col_width, col_height)
        style_shape(card, fill_color=config.COLOR_CARD_BG, line_color=config.COLOR_CARD_BORDER)

        # Red accent top bar (visual heading indicator)
        accent = slide.shapes.add_shape(1, left, CT, col_width, Inches(0.07))
        style_shape(accent, fill_color=config.COLOR_PRIMARY, line_color=None)

        # Heading inside card
        heading = col.get("heading", "")
        if heading:
            add_textbox(
                slide, _truncate(heading, 8),
                left=left + PAD, top=CT + Inches(0.12),
                width=col_width - 2 * PAD, height=Inches(0.55),
                font_size=config.CARD_HEADING_SIZE, bold=True,
                color=config.COLOR_TEXT_DARK,
            )

        # Bullet points below heading
        points = col.get("points", [])
        # Fallback: use generic message so column is never visually empty
        if not points:
            points = ["See document for details."]
        add_bullet_textbox(
            slide, [_truncate(p, 15) for p in points[:6]],
            left=left + PAD,
            top=CT + Inches(0.75),
            width=col_width - 2 * PAD,
            height=col_height - Inches(0.85),
            font_size=config.BODY_FONT_SIZE,
            color=config.COLOR_TEXT_DARK,
        )


# ---------------------------------------------------------------------------
# Layout: three_cards
# ---------------------------------------------------------------------------

def _three_cards(slide, data: dict) -> None:
    """Render a three-cards layout.

    Blueprint key expected:
        cards  (list of {number, heading, points})
    """
    cards = data.get("cards", [])[:3]
    if not cards:
        _single_focus(slide, data)
        return

    card_width = (CW - 2 * CGAP) / 3
    card_height = CH - GAP

    for i, card_data in enumerate(cards):
        left = ML + i * (card_width + CGAP)

        # Cream card with red border
        card = slide.shapes.add_shape(1, left, CT, card_width, card_height)
        style_shape(card, fill_color=config.COLOR_CARD_BG, line_color=config.COLOR_CARD_BORDER)

        # Red accent bar at top
        accent = slide.shapes.add_shape(1, left, CT, card_width, Inches(0.07))
        style_shape(accent, fill_color=config.COLOR_PRIMARY, line_color=None)

        # Large number badge
        number = card_data.get("number", str(i + 1).zfill(2))
        add_textbox(
            slide, number,
            left=left + PAD, top=CT + Inches(0.12),
            width=card_width - 2 * PAD, height=Inches(0.55),
            font_size=Pt(28), bold=True,
            color=config.COLOR_PRIMARY,
        )

        # Card heading below number
        heading = card_data.get("heading", "")
        if heading:
            add_textbox(
                slide, _truncate(heading, 7),
                left=left + PAD, top=CT + Inches(0.72),
                width=card_width - 2 * PAD, height=Inches(0.55),
                font_size=config.CARD_HEADING_SIZE, bold=True,
                color=config.COLOR_TEXT_DARK,
            )

        # Bullet points
        points = card_data.get("points", [])
        if not points:
            points = ["See document for details."]
        add_bullet_textbox(
            slide, [_truncate(p, 15) for p in points[:5]],
            left=left + PAD,
            top=CT + Inches(1.32),
            width=card_width - 2 * PAD,
            height=card_height - Inches(1.45),
            font_size=Pt(12),
            color=config.COLOR_TEXT_DARK,
        )


# ---------------------------------------------------------------------------
# Layout: key_stats
# ---------------------------------------------------------------------------

def _key_stats(slide, data: dict) -> None:
    """Render 2–4 big stat numbers with descriptive labels.

    Blueprint key expected:
        stats  (list of {value, label})
    """
    stats = data.get("stats", [])[:4]
    if not stats:
        _single_focus(slide, data)
        return

    n = len(stats)
    stat_width = (CW - (n - 1) * CGAP) / n
    stat_top = CT + Inches(0.3)
    stat_height = CH - Inches(0.6)

    for i, stat in enumerate(stats):
        left = ML + i * (stat_width + CGAP)

        # Cream card with red border
        box = slide.shapes.add_shape(1, left, stat_top, stat_width, stat_height)
        style_shape(box, fill_color=config.COLOR_CARD_BG, line_color=config.COLOR_CARD_BORDER)

        # Red accent bar at top
        accent = slide.shapes.add_shape(1, left, stat_top, stat_width, Inches(0.07))
        style_shape(accent, fill_color=config.COLOR_PRIMARY, line_color=None)

        # Big number — centered
        value = str(stat.get("value", ""))
        add_textbox(
            slide, value,
            left=left + PAD, top=stat_top + Inches(0.5),
            width=stat_width - 2 * PAD, height=Inches(1.3),
            font_size=config.STAT_NUMBER_SIZE, bold=True,
            color=config.COLOR_PRIMARY,
            align=PP_ALIGN.CENTER,
        )

        # Label below number
        label = str(stat.get("label", ""))
        add_textbox(
            slide, label,
            left=left + PAD, top=stat_top + Inches(1.9),
            width=stat_width - 2 * PAD, height=Inches(0.6),
            font_size=config.STAT_LABEL_SIZE,
            color=config.COLOR_TEXT_MUTED,
            align=PP_ALIGN.CENTER,
        )


# ---------------------------------------------------------------------------
# Layout: timeline
# ---------------------------------------------------------------------------

def _timeline(slide, data: dict) -> None:
    """Render a horizontal timeline with numbered steps.

    Blueprint key expected:
        steps  (list of {number, heading, description})
    """
    steps = data.get("steps", [])[:5]
    if not steps:
        _single_focus(slide, data)
        return

    n = len(steps)
    step_width = (CW - (n - 1) * CGAP) / n
    circle_r = Inches(0.3)
    line_y = CT + Inches(1.2)
    circle_top = line_y - circle_r
    circle_size = circle_r * 2

    # Horizontal connector line
    line = slide.shapes.add_shape(
        1,
        ML, line_y - Inches(0.02),
        CW, Inches(0.04),
    )
    style_shape(line, fill_color=config.COLOR_PRIMARY, line_color=None)

    for i, step in enumerate(steps):
        cx = ML + i * (step_width + CGAP) + step_width / 2

        # Red circle node
        circle = slide.shapes.add_shape(
            9,
            cx - circle_r, circle_top,
            circle_size, circle_size,
        )
        style_shape(circle, fill_color=config.COLOR_PRIMARY, line_color=None)

        # Step number in white inside circle
        number = step.get("number", str(i + 1).zfill(2))
        add_textbox(
            slide, number,
            left=cx - circle_r, top=circle_top,
            width=circle_size, height=circle_size,
            font_size=Pt(11), bold=True,
            color=config.COLOR_TEXT_LIGHT,
            align=PP_ALIGN.CENTER,
        )

        # Heading below circle
        heading = _truncate(step.get("heading", ""), 6)
        add_textbox(
            slide, heading,
            left=ML + i * (step_width + CGAP),
            top=line_y + circle_r + Inches(0.12),
            width=step_width, height=Inches(0.5),
            font_size=Pt(13), bold=True,
            color=config.COLOR_TEXT_DARK,
            align=PP_ALIGN.CENTER,
        )

        # Description below heading
        desc = _truncate(step.get("description", ""), 15)
        add_textbox(
            slide, desc,
            left=ML + i * (step_width + CGAP),
            top=line_y + circle_r + Inches(0.68),
            width=step_width, height=Inches(1.8),
            font_size=Pt(11),
            color=config.COLOR_TEXT_MUTED,
            align=PP_ALIGN.CENTER,
        )


# ---------------------------------------------------------------------------
# Layout: process_flow
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
    arrow_w = Inches(0.3)
    total_arrow = (n - 1) * arrow_w
    box_width = (CW - total_arrow) / n
    box_height = Inches(2.4)
    box_top = CT + (CH - box_height) / 2

    for i, step in enumerate(steps):
        left = ML + i * (box_width + arrow_w)

        box = slide.shapes.add_shape(1, left, box_top, box_width, box_height)
        style_shape(box, fill_color=config.COLOR_CARD_BG, line_color=config.COLOR_PRIMARY)

        number = step.get("number", str(i + 1).zfill(2))
        add_textbox(
            slide, number,
            left=left + PAD, top=box_top + PAD,
            width=box_width - 2 * PAD, height=Inches(0.45),
            font_size=Pt(20), bold=True, color=config.COLOR_PRIMARY,
            align=PP_ALIGN.CENTER,
        )

        heading = _truncate(step.get("heading", ""), 6)
        add_textbox(
            slide, heading,
            left=left + PAD, top=box_top + Inches(0.6),
            width=box_width - 2 * PAD, height=Inches(0.55),
            font_size=Pt(13), bold=True, color=config.COLOR_TEXT_DARK,
            align=PP_ALIGN.CENTER,
        )

        desc = _truncate(step.get("description", ""), 15)
        add_textbox(
            slide, desc,
            left=left + PAD, top=box_top + Inches(1.2),
            width=box_width - 2 * PAD, height=Inches(1.1),
            font_size=Pt(11), color=config.COLOR_TEXT_MUTED,
            align=PP_ALIGN.CENTER,
        )

        if i < n - 1:
            arrow_left = left + box_width
            arrow = slide.shapes.add_shape(
                1,
                arrow_left,
                box_top + box_height / 2 - Inches(0.04),
                arrow_w, Inches(0.08),
            )
            style_shape(arrow, fill_color=config.COLOR_PRIMARY, line_color=None)


# ---------------------------------------------------------------------------
# Layout: comparison
# ---------------------------------------------------------------------------

def _comparison(slide, data: dict) -> None:
    """Render a high-contrast side-by-side comparison with two full-color columns.

    Blueprint keys expected:
        left.heading, left.points, right.heading, right.points
    """
    col_width = (CW - CGAP) / 2
    col_height = CH - GAP
    colors = [config.COLOR_PRIMARY, config.COLOR_TEXT_DARK]

    for i, side in enumerate(("left", "right")):
        col = data.get(side, {})
        if not isinstance(col, dict):
            col = {}

        left = ML + i * (col_width + CGAP)
        bg = colors[i]

        box = slide.shapes.add_shape(1, left, CT, col_width, col_height)
        style_shape(box, fill_color=bg, line_color=None)

        text_color = config.COLOR_TEXT_LIGHT

        heading = col.get("heading", "")
        if heading:
            add_textbox(
                slide, _truncate(heading, 8),
                left=left + PAD, top=CT + PAD,
                width=col_width - 2 * PAD, height=Inches(0.55),
                font_size=config.CARD_HEADING_SIZE, bold=True,
                color=text_color, align=PP_ALIGN.CENTER,
            )

        points = col.get("points", [])
        if not points:
            points = ["See document for details."]
        add_bullet_textbox(
            slide, [_truncate(p, 15) for p in points[:6]],
            left=left + PAD, top=CT + PAD + Inches(0.65),
            width=col_width - 2 * PAD,
            height=col_height - PAD - Inches(0.75),
            font_size=config.BODY_FONT_SIZE, color=text_color,
        )


# ---------------------------------------------------------------------------
# Layout: icon_list
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

    row_height = min(CH / len(items) - CGAP, Inches(1.3))
    circle_size = Inches(0.55)
    text_left = ML + circle_size + Inches(0.25)
    text_width = CW - circle_size - Inches(0.25)

    for i, item in enumerate(items):
        row_top = CT + i * (row_height + CGAP / 2)
        circle_top = row_top + (row_height - circle_size) / 2

        # Red circle on left
        circle = slide.shapes.add_shape(9, ML, circle_top, circle_size, circle_size)
        style_shape(circle, fill_color=config.COLOR_PRIMARY, line_color=None)

        # White number inside circle
        number = item.get("number", str(i + 1).zfill(2))
        add_textbox(
            slide, number,
            left=ML, top=circle_top,
            width=circle_size, height=circle_size,
            font_size=Pt(14), bold=True, color=config.COLOR_TEXT_LIGHT,
            align=PP_ALIGN.CENTER,
        )

        # Heading to the right
        heading = _truncate(item.get("heading", ""), 8)
        add_textbox(
            slide, heading,
            left=text_left, top=row_top,
            width=text_width, height=Inches(0.4),
            font_size=config.CARD_HEADING_SIZE, bold=True,
            color=config.COLOR_TEXT_DARK,
        )

        # Description below heading
        desc = _truncate(item.get("description", ""), 20)
        add_textbox(
            slide, desc,
            left=text_left, top=row_top + Inches(0.42),
            width=text_width, height=row_height - Inches(0.45),
            font_size=config.BODY_FONT_SIZE,
            color=config.COLOR_TEXT_MUTED,
        )

        # Separator line (skip last row)
        if i < len(items) - 1:
            line = slide.shapes.add_shape(
                1,
                ML, row_top + row_height + CGAP / 4,
                CW, Inches(0.01),
            )
            style_shape(line, fill_color=config.COLOR_CARD_BORDER, line_color=None)


# ---------------------------------------------------------------------------
# Layout: single_focus
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
        add_textbox(
            slide, _truncate(focus, 20),
            left=ML, top=CT,
            width=CW, height=Inches(1.2),
            font_size=Pt(22), bold=True,
            color=config.COLOR_PRIMARY,
            align=PP_ALIGN.LEFT,
        )

        # Thin red accent line separator
        line = slide.shapes.add_shape(
            1, ML, CT + Inches(1.25), Inches(1.5), Inches(0.05)
        )
        style_shape(line, fill_color=config.COLOR_PRIMARY, line_color=None)

        bullet_top = CT + Inches(1.4)
    else:
        bullet_top = CT

    if points:
        add_bullet_textbox(
            slide, [_truncate(p, 15) for p in points[:6]],
            left=ML, top=bullet_top,
            width=CW,
            height=CH - (bullet_top - CT) - GAP,
            font_size=config.BODY_FONT_SIZE,
            color=config.COLOR_TEXT_DARK,
        )
