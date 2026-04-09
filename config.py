"""
config.py — Global constants for the MD-to-PPTX renderer.

All sizes, colors, fonts, and spacing values live here.
Renderer modules import from this file — no magic numbers elsewhere.
"""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------------------
# Slide dimensions
# Standard widescreen 16:9 format (13.33" × 7.5")
# All python-pptx positions/sizes are in EMU (English Metric Units).
# Inches() converts to EMU automatically.
# ---------------------------------------------------------------------------
SLIDE_WIDTH = Inches(13.33)   # total width of every slide
SLIDE_HEIGHT = Inches(7.5)    # total height of every slide

# ---------------------------------------------------------------------------
# Margins & safe zones
# Nothing should be placed closer than these values to any slide edge.
# This prevents text/shapes from being cut off when printed or projected.
# ---------------------------------------------------------------------------
MARGIN_LEFT = Inches(0.6)     # left safe boundary
MARGIN_RIGHT = Inches(0.6)    # right safe boundary
MARGIN_TOP = Inches(0.5)      # top safe boundary
MARGIN_BOTTOM = Inches(0.5)   # bottom safe boundary

# The usable horizontal span for content (slide width minus both side margins)
CONTENT_WIDTH = SLIDE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT   # ≈ 12.13 inches

# Where body content starts vertically — below the title area (title is ~1.2" tall + gap)
CONTENT_TOP = Inches(1.9)

# Usable vertical span for content (below title, above bottom margin)
CONTENT_HEIGHT = SLIDE_HEIGHT - CONTENT_TOP - MARGIN_BOTTOM  # ≈ 5.4 inches

# ---------------------------------------------------------------------------
# Typography — font sizes
# Hierarchy: Title > Subtitle > Card Heading > Body > Stat Label > Caption > Slide Num
# ---------------------------------------------------------------------------
TITLE_FONT_SIZE = Pt(32)        # main slide title (top of every content slide)
SUBTITLE_FONT_SIZE = Pt(18)     # subtitle on cover / section divider slides
BODY_FONT_SIZE = Pt(14)         # regular bullet point / paragraph text
CAPTION_FONT_SIZE = Pt(11)      # chart/table captions, footnotes
CARD_HEADING_SIZE = Pt(16)      # heading inside a card or column
STAT_NUMBER_SIZE = Pt(44)       # the big number in a key_stats layout
STAT_LABEL_SIZE = Pt(12)        # the label under a key_stats number
SECTION_NUMBER_SIZE = Pt(64)    # large decorative number on section divider slides
SLIDE_NUM_SIZE = Pt(10)         # slide number in the bottom-right corner

# ---------------------------------------------------------------------------
# Theme colors (Caspr / red-accent template palette)
#
# These are FALLBACK defaults used when the template's theme XML cannot be read.
# At runtime, engine.py calls config.extract_theme_colors(prs) to try to pull
# the real colors out of the loaded template.
# ---------------------------------------------------------------------------
COLOR_PRIMARY = RGBColor(0xE8, 0x3F, 0x33)      # red accent — used for headings, borders, highlights
COLOR_SECONDARY = RGBColor(0x1A, 0x1A, 0x2E)    # dark navy — alternative dark background
COLOR_TEXT_DARK = RGBColor(0x1A, 0x1A, 0x1A)    # near-black — body text on light backgrounds
COLOR_TEXT_LIGHT = RGBColor(0xFF, 0xFF, 0xFF)    # white — text on dark/colored backgrounds
COLOR_TEXT_MUTED = RGBColor(0x66, 0x66, 0x66)   # medium gray — captions, secondary labels
COLOR_CARD_BG = RGBColor(0xF9, 0xF0, 0xEE)      # light pink/cream — card background fill
COLOR_CARD_BORDER = RGBColor(0xE8, 0x3F, 0x33)  # red — card outline / border
COLOR_ALT_ROW = RGBColor(0xF5, 0xF5, 0xF5)      # light gray — alternating table row background
COLOR_HEADER_BG = RGBColor(0xE8, 0x3F, 0x33)    # red — table header row background
COLOR_DIVIDER_BG = RGBColor(0x1A, 0x1A, 0x1A)   # near-black — full-bleed section divider background

# ---------------------------------------------------------------------------
# Chart series colors — applied to chart bars/lines/slices in order.
# The list cycles if there are more series than colors.
# ---------------------------------------------------------------------------
CHART_COLORS = [
    RGBColor(0xE8, 0x3F, 0x33),  # series 1 — primary red
    RGBColor(0x1A, 0x1A, 0x1A),  # series 2 — black
    RGBColor(0xF9, 0xCB, 0xC2),  # series 3 — light pink
    RGBColor(0x66, 0x66, 0x66),  # series 4 — gray
    RGBColor(0xE8, 0x7D, 0x73),  # series 5 — medium red
    RGBColor(0xB0, 0x2A, 0x25),  # series 6 — dark red
]

# ---------------------------------------------------------------------------
# Spacing constants
# Used to ensure consistent gaps between elements across all layouts.
# ---------------------------------------------------------------------------
CARD_GAP = Inches(0.25)        # horizontal gap between side-by-side cards
ELEMENT_GAP = Inches(0.3)      # vertical gap between stacked elements
INNER_PADDING = Inches(0.2)    # padding between a card's border and its content
BULLET_INDENT = Inches(0.15)   # left indent offset for bullet point text

# ---------------------------------------------------------------------------
# Shape styling defaults
# ---------------------------------------------------------------------------
CARD_CORNER_RADIUS = Emu(60000)   # rounded corner radius for card shapes (~0.06 inches)
CARD_BORDER_WIDTH = Pt(1.5)       # thickness of card border lines
ARROW_SIZE = Inches(0.25)         # width of arrow connector shapes in process_flow layouts

# ---------------------------------------------------------------------------
# Runtime theme color extraction
# ---------------------------------------------------------------------------

def extract_theme_colors(prs) -> dict[str, RGBColor]:
    """Attempt to extract accent colors from a Presentation's theme XML.

    Reads the first slide master's theme part, finds all srgbClr elements,
    and maps known role names (accent1, dk1, lt1) to RGBColor objects.

    Falls back to module-level defaults for any color that cannot be read,
    so this function always returns a complete dict.

    Args:
        prs: A python-pptx Presentation object (already loaded).

    Returns:
        Dict with keys: "primary", "text_dark", "text_light", "card_bg".
    """
    import logging
    from lxml import etree  # lxml is bundled with python-pptx

    logger = logging.getLogger(__name__)

    # Default values to return if extraction fails
    defaults = {
        "primary": COLOR_PRIMARY,
        "text_dark": COLOR_TEXT_DARK,
        "text_light": COLOR_TEXT_LIGHT,
        "card_bg": COLOR_CARD_BG,
    }

    try:
        # Navigate to the first slide master's theme relationship
        master = prs.slide_masters[0]
        theme_part = master.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        )

        # Parse the raw XML bytes into an lxml element tree
        root = etree.fromstring(theme_part.blob)

        # The DrawingML namespace prefix for theme color elements
        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        color_map: dict[str, RGBColor] = {}

        # Walk every <a:srgbClr val="RRGGBB"> element in the theme XML
        # Each one's parent tag tells us its role (dk1, lt1, accent1, etc.)
        for elem in root.iter(f"{{{ns}}}srgbClr"):
            val = elem.get("val", "")
            if len(val) == 6:   # valid 6-digit hex color
                try:
                    r = int(val[0:2], 16)
                    g = int(val[2:4], 16)
                    b = int(val[4:6], 16)
                    # Use the parent element's local name as the role key
                    parent_tag = elem.getparent().tag.split("}")[-1]
                    color_map[parent_tag] = RGBColor(r, g, b)
                except ValueError:
                    pass   # skip any malformed hex values

        # Map standard role names to our internal color keys
        result = {
            "primary":    color_map.get("accent1", COLOR_PRIMARY),   # main accent color
            "text_dark":  color_map.get("dk1", COLOR_TEXT_DARK),      # dark text color
            "text_light": color_map.get("lt1", COLOR_TEXT_LIGHT),     # light text color
            "card_bg":    COLOR_CARD_BG,                              # no standard theme key — use default
        }
        logger.debug("Extracted theme colors: %s", result)
        return result

    except Exception as exc:
        # Any failure (missing theme part, XML parse error, etc.) → use defaults
        logger.debug("Could not extract theme colors (%s) — using defaults.", exc)
        return defaults
