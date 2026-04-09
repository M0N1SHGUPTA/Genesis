"""
agents/content_transformer.py — Agent 3: Content Transformer.

Receives the slide_plan (from StorylinePlanner) and extracted_content (from
ContentExtractor) and produces the final slide blueprint JSON — the same
schema that renderer/engine.py expects.

How it works:
  1. For each slide in the plan, look up the relevant section data from
     extracted_content and build a per-slide context dict.
  2. Send all slide contexts + a compact schema reference in one LLM call.
  3. The LLM fills in every slide's content fields (bullets, chart data, etc.)
  4. Falls back to _rule_based_transform() if the LLM fails.

The prompt is kept focused: the LLM only has to FILL IN content — it does
not have to decide layouts or slide order (already done by Agent 2).
"""

from __future__ import annotations

import json
import logging
import re

from agents.base_agent import BaseAgent

logger = logging.getLogger(__name__)

_SYSTEM = (
    "You are a presentation content writer. "
    "Fill in slide content fields based on the provided slide plan and document content. "
    "Respond with valid JSON only — no markdown, no explanation."
)

_MAX_TOKENS = 5000  # richer bullets (5-6 per slide at ~25 words) need more room


class ContentTransformer(BaseAgent):
    """Agent 3 — fills every slide's content fields from the plan + extracted data.

    Usage:
        transformer = ContentTransformer()
        blueprint = transformer.transform(slide_plan, extracted_content, parsed)
    """

    def transform(
        self,
        plan: dict,
        extracted: dict,
        parsed: dict,
    ) -> dict:
        """Produce the final slide blueprint from the plan and extracted content.

        Falls back to a rule-based transformer if the LLM fails.

        Args:
            plan:      slide_plan dict from StorylinePlanner.
            extracted: extracted_content dict from ContentExtractor.
            parsed:    original parsed dict from MarkdownParser (for fallback).

        Returns:
            Full blueprint dict consumed by renderer/engine.py.
        """
        if not self.available:
            logger.warning("ContentTransformer: LLM unavailable — using rule-based transform.")
            return self._rule_based_transform(plan, extracted, parsed)

        # Build per-slide context (maps each plan entry to its content data)
        per_slide_ctx = self._build_per_slide_context(plan, extracted)
        prompt = self._build_prompt(per_slide_ctx, extracted)

        try:
            result = self._run_with_retry(prompt, _SYSTEM, max_tokens=_MAX_TOKENS)
            if not isinstance(result.get("slides"), list) or not result["slides"]:
                raise ValueError("Transformer output missing 'slides'.")
            # Ensure title field exists at top level
            if "presentation_title" not in result:
                result["presentation_title"] = extracted.get("title", "Presentation")
            result["total_slides"] = len(result["slides"])
            logger.info(
                "ContentTransformer: %d slides filled.", result["total_slides"]
            )
            return result
        except Exception as exc:
            logger.warning("ContentTransformer failed: %s — using rule-based fallback.", exc)
            return self._rule_based_transform(plan, extracted, parsed)

    # ------------------------------------------------------------------
    # Per-slide context builder
    # ------------------------------------------------------------------

    def _build_per_slide_context(
        self,
        plan: dict,
        extracted: dict,
    ) -> list[dict]:
        """Map each slide in the plan to its source content from extracted_content.

        Args:
            plan:      slide_plan dict.
            extracted: extracted_content dict.

        Returns:
            List of per-slide context dicts, one per planned slide.
        """
        # Index key_sections by heading for fast lookup
        section_index: dict[str, dict] = {
            s["heading"]: s
            for s in extracted.get("key_sections", [])
        }

        per_slide: list[dict] = []

        for slide in plan.get("slides", []):
            slide_type = slide.get("type", "content")
            source = slide.get("source_section")

            ctx: dict = {
                "slide_number": slide["slide_number"],
                "type": slide_type,
                "layout": slide.get("layout", ""),
            }

            if slide_type == "cover":
                ctx["title"] = extracted.get("title", "")
                ctx["subtitle"] = extracted.get("subtitle", "")

            elif slide_type == "executive_summary":
                bullets = extracted.get("executive_summary_bullets", [])
                mid = max(1, len(bullets) // 2)
                ctx["title"] = "Executive Summary"
                ctx["layout"] = "two_column"
                ctx["left_points"] = bullets[:mid]
                ctx["right_points"] = bullets[mid : mid + 4] or bullets[:4]

            elif slide_type == "agenda":
                ctx["title"] = "Agenda"
                ctx["points"] = [
                    s["heading"] for s in extracted.get("key_sections", [])
                ]

            elif slide_type == "section_divider":
                sec = section_index.get(source or "", {})
                ctx["title"] = source or ""
                ctx["subtitle"] = sec.get("key_insights", [""])[0] if sec else ""

            elif slide_type == "chart":
                sec = section_index.get(source or "", {})
                ctx["title"] = source or "Data"
                ctx["chart_data"] = sec.get("chart_data", {})

            elif slide_type == "table":
                sec = section_index.get(source or "", {})
                ctx["title"] = source or "Data"
                ctx["table_data"] = sec.get("table_data", {})

            elif slide_type in ("content", "conclusion"):
                sec = section_index.get(source or "", {}) if source else {}
                ctx["title"] = source or slide_type.replace("_", " ").title()
                ctx["insights"] = sec.get("key_insights", [])
                ctx["visual_type"] = sec.get("visual_type", "none")
                # Pass along any structured data the layout might need
                if sec.get("process_steps"):
                    ctx["process_steps"] = sec["process_steps"]
                if sec.get("comparison"):
                    ctx["comparison"] = sec["comparison"]
                # For key_stats slides, pass global stats
                if slide.get("layout") == "key_stats":
                    ctx["stats"] = extracted.get("global_stats", [])
                # Conclusion: pool the best insight from every section so the
                # slide actually synthesises instead of mirroring one section.
                if slide_type == "conclusion":
                    pooled = [
                        s.get("key_insights", [""])[0]
                        for s in extracted.get("key_sections", [])
                        if s.get("key_insights")
                    ]
                    ctx["pooled_insights"] = [p for p in pooled if p][:6]
                    ctx["exec_bullets"] = extracted.get(
                        "executive_summary_bullets", []
                    )[:4]

            elif slide_type == "thank_you":
                ctx["title"] = "Thank You"
                ctx["subtitle"] = extracted.get("subtitle", "")

            per_slide.append(ctx)

        return per_slide

    # ------------------------------------------------------------------
    # Prompt builder
    # ------------------------------------------------------------------

    def _build_prompt(
        self,
        per_slide_ctx: list[dict],
        extracted: dict,
    ) -> str:
        """Build the content-filling prompt.

        Args:
            per_slide_ctx: List of per-slide context dicts.
            extracted:     Full extracted_content dict.

        Returns:
            Prompt string.
        """
        ctx_json = json.dumps(per_slide_ctx, indent=2, ensure_ascii=False)

        pres_title = extracted.get('title', 'Presentation')
        return f"""You are filling in slide content. Use ONLY the specific facts, names, and numbers from the provided slide contexts.

SLIDE CONTEXTS (each slide has pre-matched content from the document):
{ctx_json}

OUTPUT a single JSON object:
{{
  "presentation_title": "{pres_title}",
  "total_slides": N,
  "slides": [ ... one complete object per slide ... ]
}}

SLIDE SCHEMAS:

COVER: {{"slide_number": 1, "type": "cover", "title": "...", "subtitle": "ONE sentence max 20 words"}}

EXECUTIVE_SUMMARY: {{"slide_number": N, "type": "executive_summary", "layout": "two_column", "title": "Executive Summary",
  "left": {{"heading": "Key Findings", "points": ["bullet with specific fact", ...]}},
  "right": {{"heading": "Implications", "points": ["bullet with specific fact", ...]}}}}

AGENDA: {{"slide_number": N, "type": "agenda", "title": "Agenda", "points": ["Section heading 1", ...]}}

SECTION_DIVIDER: {{"slide_number": N, "type": "section_divider", "title": "Section Name", "subtitle": "one specific insight from that section"}}

CONTENT/two_column: {{"slide_number": N, "type": "content", "layout": "two_column", "title": "...",
  "left": {{"heading": "specific label", "points": ["fact-based bullet"]}},
  "right": {{"heading": "specific label", "points": ["fact-based bullet"]}}}}

CONTENT/three_cards: {{"slide_number": N, "type": "content", "layout": "three_cards", "title": "...",
  "cards": [
    {{"number": "01", "heading": "named concept", "points": ["specific detail"]}},
    {{"number": "02", "heading": "named concept", "points": ["specific detail"]}},
    {{"number": "03", "heading": "named concept", "points": ["specific detail"]}}
  ]}}

CONTENT/key_stats: {{"slide_number": N, "type": "content", "layout": "key_stats", "title": "Key Statistics",
  "stats": [{{"value": "actual number or %", "label": "what it measures"}}, ...]}}

CONTENT/timeline: {{"slide_number": N, "type": "content", "layout": "timeline", "title": "...",
  "steps": [{{"number": "01", "heading": "named step", "description": "specific detail max 12 words"}}]}}

CONTENT/process_flow: {{"slide_number": N, "type": "content", "layout": "process_flow", "title": "...",
  "steps": [{{"number": "01", "heading": "named step", "description": "specific detail max 12 words"}}]}}

CONTENT/comparison: {{"slide_number": N, "type": "content", "layout": "comparison", "title": "...",
  "left": {{"heading": "Side A name", "points": ["specific point"]}},
  "right": {{"heading": "Side B name", "points": ["specific point"]}}}}

CONTENT/icon_list: {{"slide_number": N, "type": "content", "layout": "icon_list", "title": "...",
  "items": [{{"number": "01", "heading": "named item", "description": "specific fact max 18 words"}}]}}

CONTENT/single_focus: {{"slide_number": N, "type": "content", "layout": "single_focus", "title": "...",
  "focus": "one key message with a specific fact max 20 words",
  "points": ["supporting bullet with specific detail max 15 words", ...]}}

CHART: {{"slide_number": N, "type": "chart", "title": "...", "chart_type": "bar|line|pie|area",
  "data": {{"categories": ["label1", "label2"], "series": [{{"name": "metric", "values": [10, 20]}}]}},
  "caption": "one sentence insight about the data"}}

TABLE: {{"slide_number": N, "type": "table", "title": "...",
  "table": {{"headers": ["Col1", "Col2"], "rows": [["val", "val"]]}},
  "caption": "optional one-sentence insight"}}

CONCLUSION: {{"slide_number": N, "type": "conclusion", "layout": "single_focus", "title": "Key Takeaways",
  "focus": "one sentence summarising the main strategic conclusion",
  "points": ["specific takeaway 1", "specific takeaway 2", "specific takeaway 3", "specific takeaway 4"]}}

THANK_YOU: {{"slide_number": N, "type": "thank_you", "title": "Thank You", "subtitle": "brief closing remark"}}

STRICT CONTENT RULES — these are non-negotiable:
1. Every bullet, focus, description, and stat MUST come from the insights/data in the slide contexts above. Do NOT invent facts.
2. NEVER repeat the document subtitle or cover text on any content slide.
3. NEVER use vague filler like "see document", "various factors", "strategic objectives", "key aspects", or "further analysis required".
4. Each slide must have UNIQUE content — the same sentence must not appear on two different slides.
5. Bullet points MUST be 15–22 words each and contain at least one specific name, number, date, percentage, or dollar figure. Never one-word bullets.
6. Every content slide MUST produce enough bullets to fill its layout:
     - two_column / comparison / executive_summary: 3 points per side (6 total)
     - three_cards: 3 cards, each with 2 bullet points that expand the card heading
     - icon_list: 4 items, each with a 15–22 word description
     - timeline / process_flow: 4–5 steps, each description 12–18 words with a concrete detail
     - single_focus: 4–5 supporting points
     - key_stats: 3–4 stats using actual figures
7. Cover subtitle: max 20 words, one sentence.
8. Chart values MUST be plain numbers (not strings): [10, 25, 47] not ["10", "25", "47"].
9. three_cards: exactly 3 cards with distinct, specific headings (not "Point 1 / Point 2 / Point 3").
10. key_stats values must be actual figures from the document (e.g. "$69.7B", "80+", "35%"), never "—" or "N/A".
11. conclusion: the focus must synthesise the document's main thesis; the 4 points must each draw on DIFFERENT sections (use pooled_insights when provided).
12. If a slide's insights list contains fewer than 3 items, expand by rewording paragraphs from the same section — never pad with placeholders.
"""

    # ------------------------------------------------------------------
    # Rule-based fallback transform
    # ------------------------------------------------------------------

    def _rule_based_transform(
        self,
        plan: dict,
        extracted: dict,
        parsed: dict,
    ) -> dict:
        """Build a full blueprint from the plan + extracted content without the LLM.

        This is a Python-only fallback that maps the plan to content fields
        using deterministic rules. It produces a valid blueprint even when
        the Groq API is completely unavailable.

        Args:
            plan:      slide_plan from StorylinePlanner.
            extracted: extracted_content from ContentExtractor.
            parsed:    original parsed dict (used for fallback text).

        Returns:
            Full blueprint dict.
        """
        logger.info("ContentTransformer: building rule-based blueprint from plan.")

        section_index: dict[str, dict] = {
            s["heading"]: s
            for s in extracted.get("key_sections", [])
        }

        slides: list[dict] = []

        for planned in plan.get("slides", []):
            slide_type = planned.get("type", "content")
            layout = planned.get("layout", "single_focus")
            source = planned.get("source_section")
            sec = section_index.get(source or "", {}) if source else {}
            num = planned["slide_number"]

            try:
                slide = self._build_slide(
                    slide_type, layout, num, sec, source,
                    extracted, parsed,
                )
                slides.append(slide)
            except Exception as exc:
                logger.warning("Rule-based slide %d failed: %s — using placeholder.", num, exc)
                slides.append({
                    "slide_number": num,
                    "type": "content",
                    "layout": "single_focus",
                    "title": source or "Slide",
                    "focus": "",
                    "points": [],
                })

        return {
            "presentation_title": extracted.get("title", parsed.get("title", "Presentation")),
            "total_slides": len(slides),
            "slides": slides,
        }

    def _build_slide(
        self,
        slide_type: str,
        layout: str,
        num: int,
        sec: dict,
        source: str | None,
        extracted: dict,
        parsed: dict,
    ) -> dict:
        """Build a single slide dict from rule-based logic.

        Args:
            slide_type: Slide type string (cover, content, chart, etc.)
            layout:     Layout name for content slides.
            num:        Slide number.
            sec:        Matching key_section dict from extracted content.
            source:     Source section heading string.
            extracted:  Full extracted_content dict.
            parsed:     Original parsed dict.

        Returns:
            Single slide dict.
        """
        insights: list[str] = [
            ins for ins in sec.get("key_insights", []) if ins and ins.strip()
        ]

        # Pool the top insight from every other section as a top-up source
        # when the current section doesn't carry enough bullets.
        pooled_backup: list[str] = []
        for other in extracted.get("key_sections", []):
            if other.get("heading") == sec.get("heading"):
                continue
            for ins in other.get("key_insights", [])[:2]:
                if ins and ins not in insights and ins not in pooled_backup:
                    pooled_backup.append(ins)

        def _ensure(n: int, src: list[str]) -> list[str]:
            """Return at least n items by topping up from pooled_backup."""
            out = list(src)
            i = 0
            while len(out) < n and i < len(pooled_backup):
                if pooled_backup[i] not in out:
                    out.append(pooled_backup[i])
                i += 1
            return out

        if slide_type == "cover":
            subtitle = extracted.get("subtitle", "")
            words = subtitle.split()
            if len(words) > 20:
                subtitle = " ".join(words[:20]) + "…"
            return {
                "slide_number": num,
                "type": "cover",
                "title": extracted.get("title", parsed.get("title", "Presentation")),
                "subtitle": subtitle,
            }

        if slide_type == "executive_summary":
            bullets = extracted.get("executive_summary_bullets", [])
            bullets = [b for b in bullets if b and b.strip()]
            if len(bullets) < 4:
                bullets = _ensure(4, bullets)
            mid = max(1, len(bullets) // 2)
            return {
                "slide_number": num,
                "type": "executive_summary",
                "layout": "two_column",
                "title": "Executive Summary",
                "left": {
                    "heading": "Key Findings",
                    "points": bullets[:mid] or bullets[:1],
                },
                "right": {
                    "heading": "Implications",
                    "points": bullets[mid : mid + 4] or bullets[-3:],
                },
            }

        if slide_type == "agenda":
            headings = [s["heading"] for s in extracted.get("key_sections", [])]
            return {
                "slide_number": num,
                "type": "agenda",
                "title": "Agenda",
                "points": headings[:12] or ["Overview"],
            }

        if slide_type == "section_divider":
            return {
                "slide_number": num,
                "type": "section_divider",
                "title": source or "Section",
                "subtitle": insights[0] if insights else "",
            }

        if slide_type == "thank_you":
            return {
                "slide_number": num,
                "type": "thank_you",
                "title": "Thank You",
                "subtitle": extracted.get("subtitle", ""),
            }

        if slide_type == "chart":
            chart_data = sec.get("chart_data", {})
            if not chart_data:
                chart_data = {
                    "chart_type": "bar",
                    "categories": ["A", "B", "C"],
                    "series": [{"name": "Value", "values": [1, 2, 3]}],
                }
            return {
                "slide_number": num,
                "type": "chart",
                "title": source or "Data",
                "chart_type": chart_data.get("chart_type", "bar"),
                "data": {
                    "categories": chart_data.get("categories", []),
                    "series": chart_data.get("series", []),
                },
                "caption": insights[0] if insights else "",
            }

        if slide_type == "table":
            table_data = sec.get("table_data", {"headers": ["Column"], "rows": []})
            return {
                "slide_number": num,
                "type": "table",
                "title": source or "Data",
                "table": table_data,
                "caption": insights[0] if insights else "",
            }

        # Conclusion gets pooled content from every section, not just one
        if slide_type == "conclusion":
            pooled: list[str] = []
            for s in extracted.get("key_sections", []):
                if s.get("key_insights"):
                    pooled.append(s["key_insights"][0])
            pooled = [p for p in pooled if p][:6]
            focus = (
                extracted.get("executive_summary_bullets", [""])[0]
                or (pooled[0] if pooled else "")
                or "Key strategic takeaways from the analysis"
            )
            return {
                "slide_number": num,
                "type": "conclusion",
                "layout": "single_focus",
                "title": "Key Takeaways",
                "focus": focus,
                "points": pooled[:5] or insights[:5],
            }

        # Content slides — dispatch by layout
        title = source or layout.replace("_", " ").title()

        if layout == "three_cards":
            # Each card gets one primary insight + one distinct extra.
            topped = _ensure(3, insights)
            extras_pool = [x for x in topped[3:]] + [
                x for x in pooled_backup if x not in topped
            ]
            extras_iter = iter(extras_pool)

            cards: list[dict] = []
            for i in range(3):
                if i < len(topped):
                    primary = topped[i]
                    heading = _first_words(primary, 6).rstrip(".,;:")
                    points = [primary]
                    # Each card gets a UNIQUE secondary bullet
                    try:
                        extra = next(extras_iter)
                        if extra and extra not in points:
                            points.append(extra)
                    except StopIteration:
                        pass
                else:
                    heading = f"Point {i + 1}"
                    points = [title]
                cards.append({
                    "number": str(i + 1).zfill(2),
                    "heading": heading,
                    "points": points,
                })
            return {
                "slide_number": num,
                "type": "content",
                "layout": "three_cards",
                "title": title,
                "cards": cards,
            }

        if layout in ("two_column", "comparison"):
            topped = _ensure(6, insights)
            comp = sec.get("comparison")
            if comp:
                left_pts = comp.get("left_points") or topped[:3]
                right_pts = comp.get("right_points") or topped[3:6]
                left_h = comp.get("left_heading", "Overview")
                right_h = comp.get("right_heading", "Details")
            else:
                mid = max(1, min(3, len(topped) // 2 or 1))
                left_pts = topped[:mid]
                right_pts = topped[mid : mid + 3] or topped[:3]
                left_h = "Highlights"
                right_h = "Details"
            return {
                "slide_number": num,
                "type": "content",
                "layout": layout,
                "title": title,
                "left": {"heading": left_h, "points": left_pts[:4]},
                "right": {"heading": right_h, "points": right_pts[:4]},
            }

        if layout == "key_stats":
            stats = [
                s for s in extracted.get("global_stats", [])
                if s.get("value") and s.get("value") != "—"
            ]
            if len(stats) < 2:
                # Derive stats from numerical data found in sections
                for s in extracted.get("key_sections", []):
                    cd = s.get("chart_data") or {}
                    series = (cd.get("series") or [{}])[0]
                    cats = cd.get("categories", [])
                    vals = series.get("values", [])
                    for cat, val in zip(cats, vals):
                        if len(stats) >= 4:
                            break
                        stats.append({"value": str(val), "label": f"{s.get('heading','')[:30]} {cat}".strip()})
                    if len(stats) >= 4:
                        break
            if not stats:
                stats = [{"value": "—", "label": "Key metric"}]
            return {
                "slide_number": num,
                "type": "content",
                "layout": "key_stats",
                "title": "Key Statistics",
                "stats": stats[:4],
            }

        if layout in ("timeline", "process_flow"):
            topped = _ensure(4, insights)
            steps = sec.get("process_steps") or [
                {
                    "number": str(i + 1).zfill(2),
                    "heading": _first_words(ins, 5).rstrip(".,;:"),
                    "description": ins,
                }
                for i, ins in enumerate(topped[:5])
            ]
            return {
                "slide_number": num,
                "type": "content",
                "layout": layout,
                "title": title,
                "steps": steps[:5],
            }

        if layout == "icon_list":
            topped = _ensure(4, insights)
            items = [
                {
                    "number": str(i + 1).zfill(2),
                    "heading": _first_words(ins, 5).rstrip(".,;:"),
                    "description": ins,
                }
                for i, ins in enumerate(topped[:4])
            ]
            return {
                "slide_number": num,
                "type": "content",
                "layout": "icon_list",
                "title": title,
                "items": items,
            }

        # single_focus (catch-all)
        topped = _ensure(5, insights)
        focus = topped[0] if topped else title
        return {
            "slide_number": num,
            "type": "content",
            "layout": "single_focus",
            "title": title,
            "focus": focus,
            "points": topped[1:6] if len(topped) > 1 else topped,
        }


# ---------------------------------------------------------------------------
# Module-level helpers
# ---------------------------------------------------------------------------

def _first_words(text: str, n: int) -> str:
    """Return the first n words of text, appending '…' if trimmed.

    Used only for generating short headings from a long sentence — never
    for bullet content itself.
    """
    if not text:
        return ""
    words = text.split()
    if len(words) <= n:
        return text.strip()
    return " ".join(words[:n]) + "…"
