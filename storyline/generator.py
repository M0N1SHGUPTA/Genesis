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


# ---------------------------------------------------------------------------
# Module-level text helpers (used by fallback blueprint builder)
# ---------------------------------------------------------------------------

def _first_words(text: str, n: int) -> str:
    """Return the first n words of text, appending '…' if trimmed."""
    words = text.split()
    if len(words) <= n:
        return text.strip()
    return " ".join(words[:n]) + "…"


def _extract_bullets(text: str, max_bullets: int = 6, max_words: int = 15) -> list[str]:
    """Split a block of text into concise bullet-sized strings.

    Tries sentence boundaries first; falls back to word-count chunks.
    Each returned string is at most max_words words.
    """
    if not text:
        return []
    # Split on sentence-ending punctuation followed by whitespace
    sentences = re.split(r"(?<=[.!?])\s+", text.strip())
    bullets: list[str] = []
    for sent in sentences:
        sent = sent.strip()
        if not sent:
            continue
        words = sent.split()
        if len(words) <= max_words:
            bullets.append(sent)
        else:
            # Long sentence: break into max_words chunks
            for start in range(0, len(words), max_words):
                chunk = " ".join(words[start : start + max_words])
                if chunk:
                    bullets.append(chunk)
        if len(bullets) >= max_bullets:
            break
    return bullets[:max_bullets]

# Load .env from the project root (wherever main.py is run from)
load_dotenv()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
MODEL = "llama-3.3-70b-versatile"
MAX_RETRIES = 3
# Free Groq tier allows ~12k TPM; ~1 token ≈ 4 chars → 12000 tokens ≈ 48000 chars.
# Use condensed summary prompt for anything larger to avoid 413 errors.
FULL_PROMPT_CHAR_LIMIT = 24_000
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
        # Build a condensed fallback prompt in case the full prompt is too large
        condensed_prompt = build_summary_prompt(parsed, self.target_slides)

        for attempt in range(1, MAX_RETRIES + 1):
            logger.info("LLM attempt %d/%d …", attempt, MAX_RETRIES)
            try:
                if attempt == 1:
                    raw_response = self._call_llm(prompt)
                elif "413" in last_error or "too large" in last_error.lower() or "rate_limit" in last_error.lower():
                    # Payload too large or rate limit — retry with condensed prompt
                    logger.info("Retrying with condensed summary prompt …")
                    raw_response = self._call_llm(condensed_prompt)
                else:
                    # JSON was malformed — send correction prompt referencing bad output
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
                max_tokens=3500,   # blueprint JSON needs ~2k tokens; 8192 blows the TPM limit
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
        """Validate and auto-repair the blueprint dict.

        Fixes common LLM omissions in-place instead of raising. Only raises
        if the slides list is missing or empty (unrecoverable).

        Args:
            blueprint: The parsed blueprint dict.

        Raises:
            ValueError: If slides list is absent or empty.
        """
        slides = blueprint.get("slides", [])
        if not isinstance(slides, list) or len(slides) == 0:
            raise ValueError("Blueprint has no slides.")

        # Auto-repair top-level missing keys
        if "presentation_title" not in blueprint:
            blueprint["presentation_title"] = "Presentation"
            logger.warning("Auto-repaired: missing presentation_title.")

        # Fix missing / wrong slide_number on every slide
        for i, slide in enumerate(slides):
            if "slide_number" not in slide:
                slide["slide_number"] = i + 1
                logger.debug("Auto-repaired slide_number for slide index %d.", i)

            # Fix invalid or missing type
            slide_type = slide.get("type", "")
            if slide_type not in _VALID_TYPES:
                slide["type"] = "content"
                logger.warning(
                    "Slide %d had invalid type '%s' — set to 'content'.", i + 1, slide_type
                )

        # Ensure first slide is cover
        if slides[0].get("type") != "cover":
            logger.warning("First slide is not 'cover' — injecting cover slide.")
            slides.insert(0, {
                "slide_number": 1,
                "type": "cover",
                "title": blueprint.get("presentation_title", "Presentation"),
                "subtitle": "",
            })
            for i, s in enumerate(slides):
                s["slide_number"] = i + 1

        # Ensure last slide is thank_you
        if slides[-1].get("type") != "thank_you":
            logger.warning("Last slide is not 'thank_you' — appending.")
            slides.append({
                "slide_number": len(slides) + 1,
                "type": "thank_you",
                "title": "Thank You",
                "subtitle": "",
            })

        # Fix slide count
        blueprint["total_slides"] = len(slides)

        # Clamp to 10-15 range — log but don't fail
        total = blueprint["total_slides"]
        if not (MIN_SLIDES <= total <= MAX_SLIDES):
            logger.warning(
                "total_slides=%d is outside 10–15 range — proceeding anyway.", total
            )

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

        # Slide 1: Cover — subtitle is a single short sentence (max 20 words)
        raw_subtitle = parsed.get("subtitle", "")
        subtitle_words = raw_subtitle.split()
        cover_subtitle = " ".join(subtitle_words[:20]) + ("…" if len(subtitle_words) > 20 else "")
        slides.append({
            "slide_number": 1,
            "type": "cover",
            "title": parsed.get("title", "Presentation"),
            "subtitle": cover_subtitle,
        })

        # Slide 2: Executive summary (if present)
        exec_sum = parsed.get("executive_summary", "")
        if exec_sum:
            bullets = _extract_bullets(exec_sum, max_bullets=8, max_words=14)
            # Ensure at least 2 bullets per side
            if len(bullets) < 4:
                bullets = bullets + [_first_words(exec_sum, 14)] * (4 - len(bullets))
            mid = max(2, len(bullets) // 2)
            slides.append({
                "slide_number": len(slides) + 1,
                "type": "executive_summary",
                "layout": "two_column",
                "title": "Executive Summary",
                "left": {
                    "heading": "Key Findings",
                    "points": bullets[:mid],
                },
                "right": {
                    "heading": "Implications",
                    "points": bullets[mid : mid + 4] or bullets[:4],
                },
            })

        # Slide 3: Agenda — list all section headings
        sections = parsed.get("sections", [])
        if sections:
            slides.append({
                "slide_number": len(slides) + 1,
                "type": "agenda",
                "title": "Agenda",
                "points": [sec["heading"] for sec in sections[:12]],
            })

        layout_cycle = ["three_cards", "two_column", "key_stats", "icon_list", "timeline", "two_column", "single_focus"]

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

            # Gather bullets from all subsections (with 15-word cap each)
            all_bullets: list[str] = []
            for sub in sec.get("subsections", []):
                for b in sub.get("bullets", []):
                    all_bullets.append(_first_words(b, 15))
                for para in sub.get("paragraphs", []):
                    all_bullets.extend(_extract_bullets(para, max_bullets=2, max_words=15))

            # Fallback: derive bullets from section-level content text
            if not all_bullets and sec.get("content"):
                all_bullets = _extract_bullets(sec["content"], max_bullets=6, max_words=15)

            # Last resort: use subsection headings as bullets
            if not all_bullets:
                all_bullets = [
                    _first_words(sub["heading"], 15)
                    for sub in sec.get("subsections", [])
                    if sub.get("heading")
                ][:6] or [sec["heading"]]

            layout = layout_cycle[idx % len(layout_cycle)]

            # Build slide according to layout type
            if layout == "three_cards":
                # Pick 3 subsections or split bullets into 3 groups
                subs = sec.get("subsections", [])
                if len(subs) >= 3:
                    cards = [
                        {
                            "number": str(j + 1).zfill(2),
                            "heading": _first_words(subs[j]["heading"], 5),
                            "points": ([_first_words(b, 15) for b in subs[j].get("bullets", [])]
                                       or _extract_bullets(subs[j].get("paragraphs", [""])[0] if subs[j].get("paragraphs") else "", 3, 15)
                                       or [_first_words(subs[j]["heading"], 15)]),
                        }
                        for j in range(3)
                    ]
                else:
                    chunk = max(1, len(all_bullets) // 3)
                    cards = [
                        {
                            "number": str(j + 1).zfill(2),
                            "heading": (
                                _first_words(all_bullets[j * chunk], 5)
                                if all_bullets and j * chunk < len(all_bullets)
                                else f"Point {j + 1}"
                            ),
                            "points": all_bullets[j * chunk:(j + 1) * chunk][:4] or ["See document for details."],
                        }
                        for j in range(3)
                    ]
                slide: dict = {
                    "slide_number": len(slides) + 1,
                    "type": "content",
                    "layout": "three_cards",
                    "title": sec["heading"],
                    "cards": cards,
                }

            elif layout in ("two_column", "key_stats") and all_bullets:
                mid = max(1, len(all_bullets) // 2)
                slide = {
                    "slide_number": len(slides) + 1,
                    "type": "content",
                    "layout": "two_column",
                    "title": sec["heading"],
                    "left": {
                        "heading": "Overview",
                        "points": all_bullets[:mid][:4],
                    },
                    "right": {
                        "heading": "Details",
                        "points": all_bullets[mid : mid + 4] or all_bullets[:4],
                    },
                }

            elif layout == "icon_list":
                subs = sec.get("subsections", [])[:4]
                items = [
                    {
                        "number": str(j + 1).zfill(2),
                        "heading": _first_words(sub["heading"], 6),
                        "description": (
                            _first_words(sub.get("bullets", [""])[0], 20)
                            if sub.get("bullets")
                            else _first_words(sub.get("paragraphs", [""])[0], 20)
                            if sub.get("paragraphs")
                            else sec["heading"]
                        ),
                    }
                    for j, sub in enumerate(subs)
                ]
                if not items:
                    items = [
                        {"number": str(j + 1).zfill(2), "heading": _first_words(b, 6), "description": b}
                        for j, b in enumerate(all_bullets[:4])
                    ]
                slide = {
                    "slide_number": len(slides) + 1,
                    "type": "content",
                    "layout": "icon_list",
                    "title": sec["heading"],
                    "items": items or [{"number": "01", "heading": sec["heading"], "description": sec.get("content", "")[:100]}],
                }

            else:
                slide = {
                    "slide_number": len(slides) + 1,
                    "type": "content",
                    "layout": "single_focus",
                    "title": sec["heading"],
                    "focus": _first_words(sec.get("content", sec["heading"]), 18),
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
                        # Skip chart if fewer than 2 data points — single bar is meaningless
                        if len(categories) >= 2:
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
