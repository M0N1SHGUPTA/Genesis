"""
generator.py — Stage 2: structured parsed dict → slide blueprint JSON.

Uses the Groq API (llama-3.3-70b-versatile) to intelligently decide:
  - Slide count (10–15)
  - Layout type per slide
  - Which data becomes a chart vs. text vs. table
  - Content trimming and narrative flow

Reads GROQ_API_KEY from the .env file in the project root.
"""

from __future__ import annotations

import json
import logging
import os
import re

from dotenv import load_dotenv
from groq import Groq

from storyline.prompts import (
    build_blueprint_prompt,
    build_correction_prompt,
    build_summary_prompt,
)

logger = logging.getLogger(__name__)

# Load .env from the project root (wherever main.py is run from)
load_dotenv()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
MODEL = "llama-3.3-70b-versatile"
MAX_RETRIES = 3
# Rough character limit before we switch to the condensed summary prompt.
# llama-3.3-70b-versatile has a 128k token context; ~400k chars is a safe cap.
FULL_PROMPT_CHAR_LIMIT = 80_000
MIN_SLIDES = 10
MAX_SLIDES = 15

# Required top-level keys in a valid blueprint
_REQUIRED_KEYS = {"presentation_title", "total_slides", "slides"}
# Required keys in every slide object
_REQUIRED_SLIDE_KEYS = {"slide_number", "type"}
# All valid slide type strings
_VALID_TYPES = {
    "cover", "agenda", "executive_summary", "section_divider",
    "content", "chart", "table", "conclusion", "thank_you",
}


# ---------------------------------------------------------------------------
# Public class
# ---------------------------------------------------------------------------

class StorylineGenerator:
    """Generate a slide blueprint from parsed markdown content using an LLM.

    Usage:
        generator = StorylineGenerator(target_slides=12)
        blueprint = generator.generate(parsed_dict)
    """

    def __init__(self, target_slides: int | None = None) -> None:
        """Initialise the generator.

        Args:
            target_slides: Desired number of slides (10–15), or None to let
                           the LLM decide based on content density.
        """
        self.target_slides = target_slides
        self._client = self._init_client()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def generate(self, parsed: dict) -> dict:
        """Generate a slide blueprint from a parsed markdown document.

        Tries the LLM up to MAX_RETRIES times. Falls back to a rule-based
        generator if all LLM attempts fail.

        Args:
            parsed: Structured dict produced by Stage 1 (MarkdownParser).

        Returns:
            Slide blueprint dict matching the schema in CLAUDE.md.
        """
        if self._client is None:
            logger.warning("No Groq client — using rule-based fallback generator.")
            return self._fallback_blueprint(parsed)

        # Choose full or condensed prompt based on content size
        prompt = self._build_prompt(parsed)

        last_error = ""
        raw_response = ""

        for attempt in range(1, MAX_RETRIES + 1):
            logger.info("LLM attempt %d/%d …", attempt, MAX_RETRIES)
            try:
                if attempt == 1:
                    raw_response = self._call_llm(prompt)
                else:
                    # On retry: send the correction prompt referencing the bad output
                    correction = build_correction_prompt(
                        raw_response, last_error, self.target_slides
                    )
                    raw_response = self._call_llm(correction)

                blueprint = self._parse_json(raw_response)
                self._validate(blueprint)
                logger.info(
                    "Blueprint generated: %d slides", blueprint["total_slides"]
                )
                return blueprint

            except (ValueError, KeyError) as exc:
                last_error = str(exc)
                logger.warning("Attempt %d failed: %s", attempt, last_error)

        # All retries exhausted
        logger.error(
            "LLM failed after %d attempts — using rule-based fallback.", MAX_RETRIES
        )
        return self._fallback_blueprint(parsed)

    # ------------------------------------------------------------------
    # Groq client
    # ------------------------------------------------------------------

    def _init_client(self) -> Groq | None:
        """Initialise the Groq client from GROQ_API_KEY env variable.

        Returns None (with a warning) if the key is missing, so the caller
        can gracefully fall back to the rule-based generator.
        """
        api_key = os.getenv("GROQ_API_KEY")
        if not api_key:
            logger.warning(
                "GROQ_API_KEY not set. Add it to your .env file: GROQ_API_KEY=gsk_..."
            )
            return None
        return Groq(api_key=api_key)

    def _call_llm(self, prompt: str) -> str:
        """Send a prompt to the Groq API and return the raw text response.

        Args:
            prompt: The full prompt string.

        Returns:
            Raw string response from the LLM.

        Raises:
            ValueError: If the API call fails or returns an empty response.
        """
        try:
            response = self._client.chat.completions.create(
                model=MODEL,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are a presentation design expert. "
                            "You always respond with valid JSON only — "
                            "no markdown, no explanation, no code fences."
                        ),
                    },
                    {"role": "user", "content": prompt},
                ],
                temperature=0.3,   # low temperature for structured/deterministic output
                max_tokens=8192,
            )
            text = response.choices[0].message.content or ""
            if not text.strip():
                raise ValueError("LLM returned an empty response.")
            return text
        except Exception as exc:
            raise ValueError(f"Groq API error: {exc}") from exc

    # ------------------------------------------------------------------
    # Prompt selection
    # ------------------------------------------------------------------

    def _build_prompt(self, parsed: dict) -> str:
        """Choose full or condensed prompt based on the size of parsed content.

        Args:
            parsed: Parsed document dict from Stage 1.

        Returns:
            Prompt string to send to the LLM.
        """
        full_prompt = build_blueprint_prompt(parsed, self.target_slides)
        if len(full_prompt) <= FULL_PROMPT_CHAR_LIMIT:
            logger.debug(
                "Using full prompt (%d chars)", len(full_prompt)
            )
            return full_prompt

        logger.info(
            "Parsed content is large (%d chars) — using condensed summary prompt.",
            len(full_prompt),
        )
        return build_summary_prompt(parsed, self.target_slides)

    # ------------------------------------------------------------------
    # JSON parsing
    # ------------------------------------------------------------------

    def _parse_json(self, raw: str) -> dict:
        """Extract and parse JSON from the LLM's raw response.

        Handles common LLM quirks:
        - Markdown code fences (```json ... ```)
        - Leading/trailing whitespace
        - Text before/after the JSON object

        Args:
            raw: Raw string from the LLM.

        Returns:
            Parsed Python dict.

        Raises:
            ValueError: If no valid JSON object can be extracted.
        """
        text = raw.strip()

        # Strip markdown code fences if present
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
        text = text.strip()

        # Try direct parse first
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass

        # Find the outermost { ... } block
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            try:
                return json.loads(text[start : end + 1])
            except json.JSONDecodeError as exc:
                raise ValueError(f"JSON parse error: {exc}") from exc

        raise ValueError("No JSON object found in LLM response.")

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------

    def _validate(self, blueprint: dict) -> None:
        """Validate the blueprint dict against required schema constraints.

        Args:
            blueprint: The parsed blueprint dict.

        Raises:
            ValueError: If any constraint is violated.
        """
        # Check top-level required keys
        missing = _REQUIRED_KEYS - blueprint.keys()
        if missing:
            raise ValueError(f"Blueprint missing required keys: {missing}")

        # Validate slide count
        total = blueprint.get("total_slides", 0)
        slides = blueprint.get("slides", [])

        if not isinstance(slides, list) or len(slides) == 0:
            raise ValueError("Blueprint has no slides.")

        if not (MIN_SLIDES <= total <= MAX_SLIDES):
            raise ValueError(
                f"total_slides={total} is outside the 10–15 range."
            )

        if len(slides) != total:
            # Tolerate minor discrepancy (LLM sometimes miscounts) — just fix it
            logger.warning(
                "total_slides=%d but found %d slide objects — correcting.",
                total, len(slides),
            )
            blueprint["total_slides"] = len(slides)

        # Validate each slide
        for i, slide in enumerate(slides):
            missing_slide_keys = _REQUIRED_SLIDE_KEYS - slide.keys()
            if missing_slide_keys:
                raise ValueError(
                    f"Slide {i + 1} missing keys: {missing_slide_keys}"
                )
            slide_type = slide.get("type", "")
            if slide_type not in _VALID_TYPES:
                raise ValueError(
                    f"Slide {i + 1} has invalid type: '{slide_type}'"
                )

        # First slide must be cover
        if slides[0].get("type") != "cover":
            raise ValueError("First slide must be type 'cover'.")

        # Last slide must be thank_you
        if slides[-1].get("type") != "thank_you":
            raise ValueError("Last slide must be type 'thank_you'.")

    # ------------------------------------------------------------------
    # Rule-based fallback generator
    # ------------------------------------------------------------------

    def _fallback_blueprint(self, parsed: dict) -> dict:
        """Generate a simple slide blueprint without the LLM.

        Called when the LLM is unavailable or all retries are exhausted.
        Produces one slide per H2 section, plus mandatory cover/conclusion/thank_you.

        Args:
            parsed: Parsed document dict from Stage 1.

        Returns:
            A minimal but valid slide blueprint dict.
        """
        logger.info("Building rule-based fallback blueprint …")
        slides: list[dict] = []

        # Slide 1: Cover
        slides.append({
            "slide_number": 1,
            "type": "cover",
            "title": parsed.get("title", "Presentation"),
            "subtitle": parsed.get("subtitle", ""),
        })

        # Slide 2: Executive summary (if present)
        exec_sum = parsed.get("executive_summary", "")
        if exec_sum:
            # Split into two rough halves for two_column layout
            sentences = exec_sum.split(". ")
            mid = max(1, len(sentences) // 2)
            slides.append({
                "slide_number": len(slides) + 1,
                "type": "executive_summary",
                "layout": "two_column",
                "title": "Executive Summary",
                "left": {
                    "heading": "Overview",
                    "points": [s.strip() + "." for s in sentences[:mid] if s.strip()][:3],
                },
                "right": {
                    "heading": "Key Points",
                    "points": [s.strip() + "." for s in sentences[mid:] if s.strip()][:3],
                },
            })

        # One content slide per section
        sections = parsed.get("sections", [])
        layout_cycle = ["two_column", "single_focus", "icon_list", "two_column", "single_focus"]

        for idx, sec in enumerate(sections):
            if len(slides) >= MAX_SLIDES - 2:
                # Reserve 2 slots for conclusion + thank_you
                break

            # Section divider for every 3rd section
            if idx > 0 and idx % 3 == 0 and len(slides) < MAX_SLIDES - 3:
                slides.append({
                    "slide_number": len(slides) + 1,
                    "type": "section_divider",
                    "title": sec["heading"],
                    "subtitle": "",
                })
                if len(slides) >= MAX_SLIDES - 2:
                    break

            # Gather bullets from all subsections
            all_bullets: list[str] = []
            for sub in sec.get("subsections", []):
                all_bullets.extend(sub.get("bullets", []))
                for para in sub.get("paragraphs", []):
                    # Turn paragraphs into short bullets
                    if len(para) < 120:
                        all_bullets.append(para)

            layout = layout_cycle[idx % len(layout_cycle)]

            if layout == "two_column":
                mid = max(1, len(all_bullets) // 2)
                slide: dict = {
                    "slide_number": len(slides) + 1,
                    "type": "content",
                    "layout": "two_column",
                    "title": sec["heading"],
                    "left": {
                        "heading": "Overview",
                        "points": all_bullets[:mid][:6],
                    },
                    "right": {
                        "heading": "Details",
                        "points": all_bullets[mid : mid + 6],
                    },
                }
            else:
                slide = {
                    "slide_number": len(slides) + 1,
                    "type": "content",
                    "layout": "single_focus",
                    "title": sec["heading"],
                    "focus": sec.get("content", sec["heading"])[:80],
                    "points": all_bullets[:6],
                }

            slides.append(slide)

            # Check for chart opportunity
            for sub in sec.get("subsections", []):
                if sub.get("has_numerical_data") and len(slides) < MAX_SLIDES - 2:
                    num_data = sub["numerical_data"]
                    if num_data:
                        nd = num_data[0]
                        categories = list(nd["values"].keys())
                        values = list(nd["values"].values())
                        slides.append({
                            "slide_number": len(slides) + 1,
                            "type": "chart",
                            "title": nd.get("context", "Data"),
                            "chart_type": "bar",
                            "data": {
                                "categories": [str(c) for c in categories],
                                "series": [{"name": nd.get("context", "Values"), "values": values}],
                            },
                            "caption": "",
                        })
                    break

            # Check for table opportunity
            for sub in sec.get("subsections", []):
                if sub.get("tables") and len(slides) < MAX_SLIDES - 2:
                    tbl = sub["tables"][0]
                    slides.append({
                        "slide_number": len(slides) + 1,
                        "type": "table",
                        "title": sec["heading"],
                        "table": tbl,
                        "caption": "",
                    })
                    break

        # Pad to minimum 10 slides if needed (repeat last content slide summary)
        while len(slides) < MIN_SLIDES - 2:
            slides.append({
                "slide_number": len(slides) + 1,
                "type": "content",
                "layout": "single_focus",
                "title": parsed.get("title", "Details"),
                "focus": "Additional Information",
                "points": [],
            })

        # Conclusion
        all_section_headings = [s["heading"] for s in sections[:6]]
        slides.append({
            "slide_number": len(slides) + 1,
            "type": "conclusion",
            "layout": "single_focus",
            "title": "Key Takeaways",
            "focus": parsed.get("title", "Summary"),
            "points": all_section_headings[:6],
        })

        # Thank You — always last
        slides.append({
            "slide_number": len(slides) + 1,
            "type": "thank_you",
            "title": "Thank You",
            "subtitle": parsed.get("subtitle", ""),
        })

        # Fix slide numbers sequentially
        for i, slide in enumerate(slides):
            slide["slide_number"] = i + 1

        blueprint = {
            "presentation_title": parsed.get("title", "Presentation"),
            "total_slides": len(slides),
            "slides": slides,
        }

        logger.info("Fallback blueprint: %d slides", len(slides))
        return blueprint
