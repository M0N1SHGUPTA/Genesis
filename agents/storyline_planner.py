"""
agents/storyline_planner.py — Agent 2: Storyline Planner.

Receives the extracted_content dict from ContentExtractor and outputs a
lightweight slide_plan: the exact slide sequence with type and layout for
each slide, but WITHOUT filling in content fields.

Separating "what order / what layout" from "what content goes here" gives
each agent a narrower, more reliable task. The planner's output is small
(~500 tokens) so it rarely hits rate limits.

The slide_plan is passed to ContentTransformer (Agent 3) together with the
extracted_content dict to produce the final blueprint.
"""

from __future__ import annotations

import json
import logging

from agents.base_agent import BaseAgent

logger = logging.getLogger(__name__)

_SYSTEM = (
    "You are a presentation architect. "
    "Design the optimal slide sequence for a professional deck. "
    "Respond with valid JSON only — no markdown, no explanation."
)

_MAX_TOKENS = 1000  # slide_plan is compact; generous budget still fits easily

# Available layout options shown to the planner
_LAYOUT_REFERENCE = """
LAYOUT OPTIONS (for "content" and "executive_summary" type slides):
  "two_col_sidebar"          — red sidebar with title + content cards on right (DEFAULT for content with 4+ insights)
  "six_cards"                — 3×2 grid of small icon-cards (for sections with 5–6 insights)
  "five_cards_row"           — horizontal row of 5 cards on red background (for recap/conclusion)
  "two_column"               — left column + right column (good for comparisons, before/after)
  "three_cards"              — three equal cards in a row (good for exactly 3 concepts)
  "key_stats"                — 2–4 big KPI numbers with labels
  "timeline"                 — numbered steps on a horizontal line
  "process_flow"             — boxes with arrows (good for sequential processes)
  "comparison"               — two high-contrast colored columns
  "icon_list"                — icon circles + heading + description rows (3–4 items)
  "single_focus"             — one key message + supporting bullets
  "exec_summary_with_photo"  — red sidebar + 2x2 icon-card grid (ONLY for executive_summary)
"""


class StorylinePlanner(BaseAgent):
    """Agent 2 — designs the slide sequence and layout assignments.

    Usage:
        planner = StorylinePlanner()
        slide_plan = planner.plan(extracted_content, target_slides=12)
    """

    def plan(self, extracted: dict, target_slides: int | None = None) -> dict:
        """Design the slide sequence from extracted content.

        Falls back to a rule-based plan if the LLM is unavailable or fails.

        Args:
            extracted:     Output from ContentExtractor.extract().
            target_slides: Optional desired total slide count (10–15).
                           When None, the agent decides based on content density.

        Returns:
            slide_plan dict with "total_slides" and "slides" list.
        """
        if not self.available:
            logger.warning("StorylinePlanner: LLM unavailable — using rule-based plan.")
            return self._fallback_plan(extracted, target_slides)

        prompt = self._build_prompt(extracted, target_slides)
        try:
            result = self._run_with_retry(prompt, _SYSTEM, max_tokens=_MAX_TOKENS)
            self._repair(result, extracted)
            logger.info(
                "StorylinePlanner: %d slides planned",
                result.get("total_slides", 0),
            )
            return result
        except Exception as exc:
            logger.warning("StorylinePlanner failed: %s — using rule-based fallback.", exc)
            return self._fallback_plan(extracted, target_slides)

    # ------------------------------------------------------------------
    # Prompt builder
    # ------------------------------------------------------------------

    def _build_prompt(self, extracted: dict, target_slides: int | None) -> str:
        """Build the planning prompt from extracted content.

        Args:
            extracted:     extracted_content dict from Agent 1.
            target_slides: Optional slide count target.

        Returns:
            Full prompt string.
        """
        slide_constraint = (
            f"The deck MUST have exactly {target_slides} slides."
            if target_slides is not None
            else f"Choose between 10 and 15 slides. "
                 f"Content density suggests {extracted.get('suggested_slide_count', 12)}."
        )

        # Compact representation of key_sections for the prompt
        sections_lines: list[str] = []
        for sec in extracted.get("key_sections", []):
            vtype = sec.get("visual_type", "none")
            sections_lines.append(
                f"  - \"{sec['heading']}\" | visual_type: {vtype} | "
                f"insights: {len(sec.get('key_insights', []))}"
            )
        sections_str = "\n".join(sections_lines) or "  (no sections)"

        has_exec = bool(extracted.get("executive_summary_bullets"))
        has_stats = bool(extracted.get("global_stats"))

        # Build compact extracted JSON (sections only, no bulky data)
        compact = {
            "title": extracted.get("title", ""),
            "has_executive_summary": has_exec,
            "has_global_stats": has_stats,
            "sections": [
                {
                    "heading": s["heading"],
                    "visual_type": s.get("visual_type", "none"),
                    "insight_count": len(s.get("key_insights", [])),
                }
                for s in extracted.get("key_sections", [])
            ],
        }

        return f"""Design the slide sequence for a presentation.

CONTENT OVERVIEW:
{json.dumps(compact, indent=2)}

SECTIONS WITH VISUAL TYPES:
{sections_str}

SLIDE COUNT RULE:
{slide_constraint}

MANDATORY STRUCTURE:
1. Slide 1: type "cover"
2. Slide 2: type "executive_summary", layout "exec_summary_with_photo" — ONLY if has_executive_summary is true
3. Slide 3: type "agenda"
4. Middle slides: mix of "section_divider", "content", "chart", "table"
5. Second-to-last: type "conclusion", layout "single_focus"
6. Last slide: type "thank_you"

LAYOUT ASSIGNMENT RULES:
- "chart" type slide for any section with visual_type "chart"
- "table" type slide for any section with visual_type "table"
- "timeline" or "process_flow" layout for sections with visual_type "timeline" or "process_flow"
- "six_cards" for sections with 5–6 insights (dense info grid)
- "three_cards" layout for sections with exactly 3 insights
- "two_col_sidebar" as the DEFAULT for content slides with 4+ insights (replaces plain two_column)
- "two_column" only when visual_type is "comparison" or for before/after comparisons
- "key_stats" if has_global_stats is true — use it once for a dedicated KPI slide
- "icon_list" for sections that list 3–4 named items or recommendations
- "five_cards_row" for conclusion (red background recap with 5 key takeaways)
- "single_focus" for short sections with 1–2 insights only
- NEVER use the same layout on two consecutive "content" slides

{_LAYOUT_REFERENCE}

OUTPUT a JSON object — no explanation, no markdown:
{{
  "total_slides": 12,
  "slides": [
    {{"slide_number": 1, "type": "cover"}},
    {{"slide_number": 2, "type": "executive_summary", "layout": "exec_summary_with_photo"}},
    {{"slide_number": 3, "type": "agenda"}},
    {{"slide_number": 4, "type": "section_divider", "source_section": "Section Name"}},
    {{"slide_number": 5, "type": "content", "layout": "two_col_sidebar", "source_section": "Section Name"}},
    {{"slide_number": 6, "type": "chart", "source_section": "Section Name"}},
    {{"slide_number": 7, "type": "content", "layout": "key_stats", "source_section": null}},
    {{"slide_number": 8, "type": "conclusion", "layout": "single_focus"}},
    {{"slide_number": 9, "type": "thank_you"}}
  ]
}}

Every slide entry MUST have: slide_number, type.
Content and executive_summary slides MUST also have: layout.
Section-sourced slides (content, chart, table, section_divider) MUST have: source_section.
"""

    # ------------------------------------------------------------------
    # Repair
    # ------------------------------------------------------------------

    def _repair(self, result: dict, extracted: dict) -> None:
        """Fix common LLM omissions in the slide plan in-place.

        Args:
            result:    Parsed slide plan dict (mutated).
            extracted: Original extracted content (used for fallback values).
        """
        slides = result.get("slides", [])
        if not isinstance(slides, list) or not slides:
            raise ValueError("Slide plan has no slides — triggering fallback.")

        # Fix missing slide_numbers
        for i, slide in enumerate(slides):
            if "slide_number" not in slide:
                slide["slide_number"] = i + 1
            if "type" not in slide:
                slide["type"] = "content"

        # Ensure cover is first
        if slides[0].get("type") != "cover":
            slides.insert(0, {"slide_number": 1, "type": "cover"})

        # Ensure thank_you is last
        if slides[-1].get("type") != "thank_you":
            slides.append({"slide_number": len(slides) + 1, "type": "thank_you"})

        # Renumber sequentially
        for i, slide in enumerate(slides):
            slide["slide_number"] = i + 1

        result["total_slides"] = len(slides)

    # ------------------------------------------------------------------
    # Rule-based fallback
    # ------------------------------------------------------------------

    def _fallback_plan(
        self,
        extracted: dict,
        target_slides: int | None,
    ) -> dict:
        """Generate a rule-based slide plan without the LLM.

        Args:
            extracted:     extracted_content dict.
            target_slides: Optional target slide count.

        Returns:
            slide_plan dict.
        """
        logger.info("StorylinePlanner: building rule-based slide plan.")

        max_slides = target_slides or extracted.get("suggested_slide_count", 12)
        max_slides = max(10, min(15, max_slides))

        slides: list[dict] = []

        # 1. Cover
        slides.append({"slide_number": 1, "type": "cover"})

        # 2. Executive summary
        if extracted.get("executive_summary_bullets"):
            slides.append({
                "slide_number": 2,
                "type": "executive_summary",
                "layout": "exec_summary_with_photo",
            })

        # 3. Agenda
        slides.append({"slide_number": len(slides) + 1, "type": "agenda"})

        # 4. Global stats slide (if any)
        if extracted.get("global_stats") and len(slides) < max_slides - 3:
            slides.append({
                "slide_number": len(slides) + 1,
                "type": "content",
                "layout": "key_stats",
                "source_section": None,
            })

        _LAYOUT_CYCLE = [
            "two_col_sidebar", "three_cards", "six_cards",
            "two_col_sidebar", "icon_list", "timeline",
        ]
        layout_idx = 0

        for sec in extracted.get("key_sections", []):
            if len(slides) >= max_slides - 2:
                break

            visual = sec.get("visual_type", "none")

            # Section divider every 3 sections
            if layout_idx > 0 and layout_idx % 3 == 0 and len(slides) < max_slides - 3:
                slides.append({
                    "slide_number": len(slides) + 1,
                    "type": "section_divider",
                    "source_section": sec["heading"],
                })
                if len(slides) >= max_slides - 2:
                    break

            if visual == "chart" and len(slides) < max_slides - 2:
                slides.append({
                    "slide_number": len(slides) + 1,
                    "type": "chart",
                    "source_section": sec["heading"],
                })
            elif visual == "table" and len(slides) < max_slides - 2:
                slides.append({
                    "slide_number": len(slides) + 1,
                    "type": "table",
                    "source_section": sec["heading"],
                })
            else:
                insight_count = len(sec.get("key_insights", []))
                if insight_count in (5, 6):
                    layout = "six_cards"
                elif insight_count == 3:
                    layout = "three_cards"
                elif visual in ("timeline", "process_flow"):
                    layout = visual
                elif visual == "comparison":
                    layout = "comparison"
                else:
                    layout = _LAYOUT_CYCLE[layout_idx % len(_LAYOUT_CYCLE)]

                slides.append({
                    "slide_number": len(slides) + 1,
                    "type": "content",
                    "layout": layout,
                    "source_section": sec["heading"],
                })

            layout_idx += 1

        # Pad to minimum 10 slides
        while len(slides) < 10 - 2:  # -2 for conclusion + thank_you
            slides.append({
                "slide_number": len(slides) + 1,
                "type": "content",
                "layout": "single_focus",
                "source_section": None,
            })

        # Conclusion — five_cards_row for a punchy red recap
        slides.append({
            "slide_number": len(slides) + 1,
            "type": "conclusion",
            "layout": "five_cards_row",
        })

        # Thank you
        slides.append({
            "slide_number": len(slides) + 1,
            "type": "thank_you",
        })

        # Renumber sequentially
        for i, s in enumerate(slides):
            s["slide_number"] = i + 1

        return {"total_slides": len(slides), "slides": slides}
