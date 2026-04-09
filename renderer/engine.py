"""
renderer/engine.py — Main render loop: blueprint + template → output.pptx.

This is the Stage 3 entry point. The Renderer class:
  1. Selects the best template from the templates/ folder (or uses an explicit path)
  2. Loads the template as a Presentation object
  3. Records how many slides the template originally had
  4. Iterates over each slide in the blueprint and dispatches to the right renderer
  5. Removes the original template placeholder slides (keeping only our generated slides)
  6. Saves the final .pptx file to disk
"""

from __future__ import annotations

import logging
import os
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

import config
from renderer.utils import (
    add_textbox,
    add_slide_title,
    add_slide_number,
    get_layout_by_name,
    get_blank_layout,
    remove_template_slides,
    style_shape,
)
from renderer.layouts import render_content_slide   # handles all "content" slide types
from renderer.charts import render_chart_slide       # handles "chart" slides
from renderer.tables import render_table_slide       # handles "table" slides

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Public class
# ---------------------------------------------------------------------------

class Renderer:
    """Stage 3: Render a slide blueprint into a .pptx file.

    Instantiate once and call render() for each presentation you want to produce.

    Usage:
        renderer = Renderer(template_path=None, templates_dir="templates")
        renderer.render(blueprint, parsed, "outputs/result.pptx")
    """

    def __init__(
        self,
        template_path: str | None = None,
        templates_dir: str = "templates",
    ) -> None:
        """Store configuration; no I/O happens here.

        Args:
            template_path: Path to an explicit .pptx template file.
                           When None, the best template is auto-selected from
                           templates_dir based on content keyword matching.
            templates_dir: Folder containing available .pptx Slide Master files.
        """
        self.template_path = template_path
        self.templates_dir = templates_dir

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def render(self, blueprint: dict, parsed: dict, output_path: str) -> None:
        """Render the slide blueprint to a .pptx file.

        This is the main method. It orchestrates the entire Stage 3 process:
        select template → load presentation → add slides → remove originals → save.

        Args:
            blueprint: Slide blueprint dict produced by Stage 2 (StorylineGenerator).
                       Contains a "slides" list with one dict per slide.
            parsed: Parsed document dict from Stage 1 (MarkdownParser).
                    Used only for template auto-selection heuristics.
            output_path: File path where the finished .pptx will be saved.
        """
        # Step 1: Decide which template to use
        template = self._select_template(parsed)
        logger.info("Using template: %s", template)

        # Step 2: Load the template — this copies all Slide Master layouts into memory
        prs = Presentation(template)

        # Step 3: Record the number of original template placeholder slides.
        # We'll use this count later to remove them after adding our content slides.
        original_slide_count = len(prs.slides)

        # Step 4: Iterate over each slide in the blueprint and render it
        slides_data = blueprint.get("slides", [])
        logger.info("Rendering %d slide(s) …", len(slides_data))

        for slide_data in slides_data:
            try:
                self._render_slide(prs, slide_data)
            except Exception as exc:
                # A single slide failing should never stop the whole presentation
                logger.warning(
                    "Slide %s (%s) failed: %s — skipping.",
                    slide_data.get("slide_number", "?"),
                    slide_data.get("type", "?"),
                    exc,
                )

        # Step 5: Delete the original template placeholder slides from the front
        # (all our new slides were appended after them)
        remove_template_slides(prs, original_slide_count)

        # Step 6: Save the finished presentation to disk
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        prs.save(output_path)
        logger.info("Saved: %s", output_path)

    # ------------------------------------------------------------------
    # Template selection
    # ------------------------------------------------------------------

    def _select_template(self, parsed: dict) -> str:
        """Determine which template file to use.

        If an explicit template_path was set at construction, use it directly.
        Otherwise, scan the templates_dir for .pptx files and score each one
        by counting how many words from the document title appear in the filename.
        The highest-scoring template is selected.

        Args:
            parsed: Parsed document dict (for title keyword extraction).

        Returns:
            Absolute or relative path string to the selected .pptx template.

        Raises:
            FileNotFoundError: If no template can be found anywhere.
        """
        # If the user explicitly provided a template, just use it
        if self.template_path:
            return self.template_path

        # Scan the templates directory for all .pptx files
        templates = list(Path(self.templates_dir).glob("*.pptx"))
        if not templates:
            raise FileNotFoundError(
                f"No .pptx templates found in '{self.templates_dir}'. "
                "Add at least one Slide Master .pptx file or use --template."
            )

        # If there's only one template, no need to score — just use it
        if len(templates) == 1:
            return str(templates[0])

        # Score each template: count words from the document title (>3 chars)
        # that appear in the template's filename stem
        title = parsed.get("title", "").lower()
        best = templates[0]
        best_score = 0

        for tpl in templates:
            # Normalise the filename (remove underscores/hyphens → spaces) for word matching
            stem = tpl.stem.lower().replace("_", " ").replace("-", " ")
            score = sum(1 for word in title.split() if len(word) > 3 and word in stem)
            if score > best_score:
                best_score = score
                best = tpl

        logger.info("Auto-selected template: %s (score=%d)", best.name, best_score)
        return str(best)

    # ------------------------------------------------------------------
    # Slide dispatch
    # ------------------------------------------------------------------

    def _render_slide(self, prs: Presentation, slide_data: dict) -> None:
        """Route a single slide to the correct renderer based on its type.

        Special slide types (cover, section_divider, thank_you) have their own
        methods because they use specific template layouts and different designs.
        All other types (content, chart, table, agenda, executive_summary,
        conclusion) go through the generic content pipeline.

        Args:
            prs: The Presentation object being built.
            slide_data: Single slide dict from the blueprint JSON.
        """
        slide_type = slide_data.get("type", "content")
        slide_num = slide_data.get("slide_number", 0)

        if slide_type == "cover":
            # Uses the template's dedicated Cover layout (has background image/branding)
            self._render_cover(prs, slide_data)

        elif slide_type == "section_divider":
            # Uses the template's Divider layout for section break slides
            self._render_divider(prs, slide_data)

        elif slide_type == "thank_you":
            # Uses the template's Thank You layout for the closing slide
            self._render_thank_you(prs, slide_data)

        elif slide_type == "chart":
            # Chart slides use a Blank layout so we have full control over positioning
            layout = get_blank_layout(prs)
            slide = prs.slides.add_slide(layout)
            self._clear_placeholders(slide)   # remove any placeholder shapes from the layout
            render_chart_slide(slide, slide_data, slide_num)

        elif slide_type == "table":
            # Table slides also use Blank layout
            layout = get_blank_layout(prs)
            slide = prs.slides.add_slide(layout)
            self._clear_placeholders(slide)
            render_table_slide(slide, slide_data, slide_num)

        else:
            # All remaining types (content, agenda, executive_summary, conclusion)
            # use the Blank layout and the layouts.py dispatcher
            layout = get_blank_layout(prs)
            slide = prs.slides.add_slide(layout)
            self._clear_placeholders(slide)
            render_content_slide(slide, slide_data, slide_num)

    # ------------------------------------------------------------------
    # Specific slide type renderers
    # ------------------------------------------------------------------

    def _render_cover(self, prs: Presentation, data: dict) -> None:
        """Render the cover/title slide.

        Tries the template's Cover layout first. If it has usable placeholders,
        fills them with the title and subtitle. Otherwise falls back to
        manually positioned text boxes drawn on top of the background.
        """
        # Look for a layout named "cover" — common in most Slide Masters
        layout = get_layout_by_name(prs, "cover")
        # Some templates name the cover layout with a "1_" prefix — try that too
        if "cover" not in layout.name.lower():
            layout = get_layout_by_name(prs, "1_")

        slide = prs.slides.add_slide(layout)

        title = data.get("title", "")
        subtitle = data.get("subtitle", "")

        # Attempt to fill the template's built-in title/subtitle placeholders
        filled = self._fill_placeholders(slide, title=title, subtitle=subtitle)

        if not filled:
            # Placeholders didn't work — draw text boxes manually
            self._clear_placeholders(slide)
            # Large centred title in white (template backgrounds are usually dark)
            add_textbox(
                slide, title,
                left=config.MARGIN_LEFT,
                top=Inches(2.5),           # roughly vertically centred on a 7.5" slide
                width=config.CONTENT_WIDTH,
                height=Inches(1.5),
                font_size=Pt(40), bold=True,
                color=config.COLOR_TEXT_LIGHT,
                align=PP_ALIGN.CENTER,
            )
            if subtitle:
                add_textbox(
                    slide, subtitle,
                    left=config.MARGIN_LEFT,
                    top=Inches(4.2),       # below the title
                    width=config.CONTENT_WIDTH,
                    height=Inches(0.8),
                    font_size=config.SUBTITLE_FONT_SIZE,
                    color=config.COLOR_TEXT_LIGHT,
                    align=PP_ALIGN.CENTER,
                )

    def _render_divider(self, prs: Presentation, data: dict) -> None:
        """Render a section divider slide.

        Tries the template's Divider layout first. If its placeholders can't be
        filled, draws a full-bleed dark background with large white title text.
        """
        layout = get_layout_by_name(prs, "divider")
        slide = prs.slides.add_slide(layout)

        title = data.get("title", "")
        subtitle = data.get("subtitle", "")

        filled = self._fill_placeholders(slide, title=title, subtitle=subtitle)

        if not filled:
            self._clear_placeholders(slide)
            # Full-slide dark background to create visual contrast from content slides
            bg = slide.shapes.add_shape(
                1, 0, 0, config.SLIDE_WIDTH, config.SLIDE_HEIGHT
            )
            style_shape(bg, fill_color=config.COLOR_DIVIDER_BG, line_color=None)

            # Red accent dash as a decorative element above the title
            add_textbox(
                slide, "—",
                left=config.MARGIN_LEFT, top=Inches(2.0),
                width=Inches(1.0), height=Inches(0.6),
                font_size=Pt(48), bold=True,
                color=config.COLOR_PRIMARY,
            )
            # Large white section title
            add_textbox(
                slide, title,
                left=config.MARGIN_LEFT, top=Inches(2.7),
                width=config.CONTENT_WIDTH, height=Inches(1.5),
                font_size=Pt(36), bold=True,
                color=config.COLOR_TEXT_LIGHT,
            )
            if subtitle:
                # Muted subtitle below the title
                add_textbox(
                    slide, subtitle,
                    left=config.MARGIN_LEFT, top=Inches(4.3),
                    width=config.CONTENT_WIDTH, height=Inches(0.6),
                    font_size=config.SUBTITLE_FONT_SIZE,
                    color=config.COLOR_TEXT_MUTED,
                )

    def _render_thank_you(self, prs: Presentation, data: dict) -> None:
        """Render the closing Thank You slide.

        Tries the template's Thank You layout first, then falls back to
        a manually drawn centred layout on a default background.
        """
        # Look for a layout whose name contains "thank"
        layout = get_layout_by_name(prs, "thank")
        slide = prs.slides.add_slide(layout)

        title = data.get("title", "Thank You")
        subtitle = data.get("subtitle", "")

        filled = self._fill_placeholders(slide, title=title, subtitle=subtitle)

        if not filled:
            self._clear_placeholders(slide)
            # Large centred "Thank You" text
            add_textbox(
                slide, title,
                left=config.MARGIN_LEFT, top=Inches(2.8),
                width=config.CONTENT_WIDTH, height=Inches(1.2),
                font_size=Pt(48), bold=True,
                color=config.COLOR_TEXT_LIGHT,
                align=PP_ALIGN.CENTER,
            )
            if subtitle:
                add_textbox(
                    slide, subtitle,
                    left=config.MARGIN_LEFT, top=Inches(4.2),
                    width=config.CONTENT_WIDTH, height=Inches(0.6),
                    font_size=config.SUBTITLE_FONT_SIZE,
                    color=config.COLOR_TEXT_MUTED,
                    align=PP_ALIGN.CENTER,
                )

    # ------------------------------------------------------------------
    # Placeholder helpers
    # ------------------------------------------------------------------

    def _fill_placeholders(
        self, slide, title: str = "", subtitle: str = ""
    ) -> bool:
        """Try to populate the slide's built-in layout placeholders.

        Template layouts often have pre-positioned title (idx=0) and body/subtitle
        (idx=1) placeholders that inherit the template's font and position styling.
        Using them is preferable to drawing new text boxes from scratch.

        Args:
            slide: python-pptx slide object.
            title: Text for the title placeholder (idx=0).
            subtitle: Text for the subtitle/body placeholder (idx=1).

        Returns:
            True if at least the title placeholder was successfully populated.
            False means the caller should fall back to manual text boxes.
        """
        filled = False
        for ph in slide.placeholders:
            try:
                idx = ph.placeholder_format.idx
                if idx == 0 and title:
                    ph.text = title      # title placeholder
                    filled = True
                elif idx == 1 and subtitle:
                    ph.text = subtitle   # subtitle/body placeholder
            except Exception as exc:
                logger.debug("Placeholder fill failed (idx=%s): %s", idx, exc)
        return filled

    def _clear_placeholders(self, slide) -> None:
        """Remove all placeholder shapes from a slide.

        When we draw shapes manually (not using template placeholders), any
        existing placeholder shapes from the layout would sit underneath our
        custom shapes and could show default "Click to add title" prompt text.
        Removing them gives us a truly blank canvas.

        Args:
            slide: python-pptx slide object.
        """
        # Build a list first — modifying the collection while iterating causes issues
        placeholders = list(slide.placeholders)
        for ph in placeholders:
            try:
                sp = ph._element             # the underlying <p:sp> XML element
                sp.getparent().remove(sp)    # detach it from the slide's shape tree
            except Exception:
                pass   # if removal fails for any placeholder, just skip it
