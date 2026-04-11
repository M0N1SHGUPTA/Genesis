"""
config.py - Global constants for the MD-to-PPTX renderer.

All sizes, colors, fonts, and spacing values live here.
Renderer modules import from this file - no magic numbers elsewhere.
"""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------------------
# Slide dimensions
# Standard widescreen 16:9 format (13.33" x 7.5")
# All python-pptx positions/sizes are in EMU (English Metric Units).
# Inches() converts to EMU automatically.
# ---------------------------------------------------------------------------
SLIDE_WIDTH = Inches(13.33)
SLIDE_HEIGHT = Inches(7.5)

# ---------------------------------------------------------------------------
# Margins and safe zones
# ---------------------------------------------------------------------------
MARGIN_LEFT = Inches(0.6)
MARGIN_RIGHT = Inches(0.6)
MARGIN_TOP = Inches(0.5)
MARGIN_BOTTOM = Inches(0.5)

CONTENT_WIDTH = SLIDE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
CONTENT_TOP = Inches(1.9)
CONTENT_HEIGHT = SLIDE_HEIGHT - CONTENT_TOP - MARGIN_BOTTOM

# ---------------------------------------------------------------------------
# Typography
# ---------------------------------------------------------------------------
TITLE_FONT_SIZE = Pt(32)
SUBTITLE_FONT_SIZE = Pt(18)
BODY_FONT_SIZE = Pt(14)
CAPTION_FONT_SIZE = Pt(11)
CARD_HEADING_SIZE = Pt(16)
STAT_NUMBER_SIZE = Pt(68)
STAT_LABEL_SIZE = Pt(13)
SECTION_NUMBER_SIZE = Pt(80)
SLIDE_NUM_SIZE = Pt(10)

# ---------------------------------------------------------------------------
# Font families
# ---------------------------------------------------------------------------
TITLE_FONT = "Georgia"
BODY_FONT = None

# ---------------------------------------------------------------------------
# Theme colors
# ---------------------------------------------------------------------------
COLOR_PRIMARY = RGBColor(0xE8, 0x3F, 0x33)
COLOR_SECONDARY = RGBColor(0x1A, 0x1A, 0x2E)
COLOR_TEXT_DARK = RGBColor(0x1A, 0x1A, 0x1A)
COLOR_TEXT_LIGHT = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_TEXT_MUTED = RGBColor(0x66, 0x66, 0x66)
COLOR_TEXT_SECONDARY = RGBColor(0x1F, 0x1F, 0x1F)  # Near-black — avoids faded look on cream cards
COLOR_CARD_BG = RGBColor(0xF9, 0xF0, 0xEE)
COLOR_CARD_BORDER = RGBColor(0xE8, 0x3F, 0x33)
COLOR_ALT_ROW = RGBColor(0xF5, 0xF5, 0xF5)
COLOR_HEADER_BG = RGBColor(0xE8, 0x3F, 0x33)
COLOR_DIVIDER_BG = RGBColor(0x1A, 0x1A, 0x1A)
COLOR_ACCENT_DARK = RGBColor(0x2D, 0x2D, 0x2D)

# ---------------------------------------------------------------------------
# Chart colors
# ---------------------------------------------------------------------------
CHART_COLORS = [
    RGBColor(0xE8, 0x3F, 0x33),
    RGBColor(0x1A, 0x1A, 0x1A),
    RGBColor(0xF9, 0xCB, 0xC2),
    RGBColor(0x66, 0x66, 0x66),
    RGBColor(0xE8, 0x7D, 0x73),
    RGBColor(0xB0, 0x2A, 0x25),
]

# ---------------------------------------------------------------------------
# Spacing
# ---------------------------------------------------------------------------
CARD_GAP = Inches(0.2)
ELEMENT_GAP = Inches(0.25)
INNER_PADDING = Inches(0.2)
BULLET_INDENT = Inches(0.15)

# ---------------------------------------------------------------------------
# Shape styling defaults
# ---------------------------------------------------------------------------
CARD_CORNER_RADIUS = Emu(60000)
CARD_BORDER_WIDTH = Pt(1.5)
ARROW_SIZE = Inches(0.25)


# ---------------------------------------------------------------------------
# Runtime theme color extraction
# ---------------------------------------------------------------------------

def read_theme_color_roles(prs) -> dict[str, RGBColor]:
    """Read raw theme roles (accent1, accent2, dk1, lt1, etc.) from a template.

    PowerPoint themes can encode colors as either:
      - <a:srgbClr val="RRGGBB">
      - <a:sysClr lastClr="RRGGBB">

    Some templates in this repo use sysClr for dk1/lt1, so only reading srgbClr
    causes us to miss the real text colors and fall back to defaults.
    """
    from lxml import etree

    try:
        master = prs.slide_masters[0]
        theme_part = master.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        )
        root = etree.fromstring(theme_part.blob)
        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        scheme = root.find(f".//{{{ns}}}clrScheme")
        if scheme is None:
            return {}

        roles: dict[str, RGBColor] = {}
        for child in scheme:
            role = child.tag.split("}")[-1]

            srgb = child.find(f"{{{ns}}}srgbClr")
            if srgb is not None:
                hex_color = srgb.get("val", "")
            else:
                sys_clr = child.find(f"{{{ns}}}sysClr")
                hex_color = sys_clr.get("lastClr", "") if sys_clr is not None else ""

            if len(hex_color) != 6:
                continue

            try:
                roles[role] = RGBColor(
                    int(hex_color[0:2], 16),
                    int(hex_color[2:4], 16),
                    int(hex_color[4:6], 16),
                )
            except ValueError:
                continue

        return roles

    except Exception:
        return {}


def extract_theme_colors(prs) -> dict[str, RGBColor]:
    """Extract the most useful template colors with safe fallbacks."""
    import logging

    logger = logging.getLogger(__name__)

    defaults = {
        "primary": COLOR_PRIMARY,
        "text_dark": COLOR_TEXT_DARK,
        "text_light": COLOR_TEXT_LIGHT,
        "card_bg": COLOR_CARD_BG,
    }

    try:
        color_map = read_theme_color_roles(prs)
        raw_accent1 = color_map.get("accent1")
        accent2 = color_map.get("accent2")

        if raw_accent1 and (raw_accent1[0] + raw_accent1[1] + raw_accent1[2]) / 3 > 200:
            primary = accent2 or raw_accent1
        else:
            primary = raw_accent1 or COLOR_PRIMARY

        result = {
            "primary": primary,
            "text_dark": color_map.get("dk1", COLOR_TEXT_DARK),
            "text_light": color_map.get("lt1", COLOR_TEXT_LIGHT),
            "card_bg": COLOR_CARD_BG,
        }
        logger.debug("Extracted theme colors: %s", result)
        return result

    except Exception as exc:
        logger.debug("Could not extract theme colors (%s) - using defaults.", exc)
        return defaults
