"""
agents/content_extractor.py — Agent 1: Content Extractor.

Receives the full structured parsed dict from Stage 1 (MarkdownParser) and
produces a leaner, semantically enriched extracted_content dict.

What this agent adds over the raw parsed dict:
  - Prioritises the 3-4 most important insights per section (not a raw dump)
  - Classifies each section's best visual type (chart / table / timeline / etc.)
  - Extracts ready-to-use chart / table / process data for each section
  - Identifies global statistics for key_stats slides
  - Suggests a sensible slide count based on content density

The output is passed directly to StorylinePlanner (Agent 2).
"""

from __future__ import annotations

import json
import logging

from agents.base_agent import BaseAgent

logger = logging.getLogger(__name__)

_SYSTEM = (
    "You are a content analyst specialising in presentation design. "
    "Your job is to extract the most important insights and visual opportunities "
    "from documents. Respond with valid JSON only — no markdown, no explanation."
)

# Token budget: richer insights need more room. 3500 comfortably fits
# 10+ sections × 6 insights × ~25 words.
_MAX_TOKENS = 3500


class ContentExtractor(BaseAgent):
    """Agent 1 — distils a parsed document dict into key content for the planner.

    Usage:
        extractor = ContentExtractor()
        extracted = extractor.extract(parsed)
    """

    def extract(self, parsed: dict) -> dict:
        """Extract key content, visual candidates, and global stats from parsed dict.

        Falls back to a rule-based extraction if the LLM is unavailable or fails.

        Args:
            parsed: Structured dict from MarkdownParser.parse().

        Returns:
            extracted_content dict consumed by StorylinePlanner.
        """
        if not self.available:
            logger.warning("ContentExtractor: LLM unavailable — using rule-based extraction.")
            return self._fallback_extract(parsed)

        prompt = self._build_prompt(parsed)
        try:
            result = self._run_with_retry(prompt, _SYSTEM, max_tokens=_MAX_TOKENS)
            self._repair(result, parsed)
            logger.info(
                "ContentExtractor: %d key section(s), %d global stat(s), suggested %d slides",
                len(result.get("key_sections", [])),
                len(result.get("global_stats", [])),
                result.get("suggested_slide_count", 0),
            )
            return result
        except Exception as exc:
            logger.warning("ContentExtractor failed: %s — using rule-based fallback.", exc)
            return self._fallback_extract(parsed)

    # ------------------------------------------------------------------
    # Prompt builder
    # ------------------------------------------------------------------

    def _build_prompt(self, parsed: dict) -> str:
        """Build the condensed extraction prompt.

        Sends a generous, content-rich view of the parsed document so the LLM
        has enough real facts (company names, numbers, dates, dollar figures)
        to produce specific, non-generic insights. We deliberately do NOT
        pre-summarise here — the LLM's job is to distil.
        """
        lines: list[str] = [
            f"Title: {parsed.get('title', 'Untitled')}",
            f"Subtitle: {parsed.get('subtitle', '')[:300]}",
        ]

        exec_sum = parsed.get("executive_summary", "")
        if exec_sum:
            lines.append(f"\nExecutive Summary:\n{exec_sum[:2000]}")

        for sec in parsed.get("sections", []):
            lines.append(f"\n## {sec['heading']}")
            if sec.get("content"):
                lines.append(f"  {sec['content'][:500]}")

            for sub in sec.get("subsections", []):
                if sub.get("heading"):
                    lines.append(f"  ### {sub['heading']}")
                # Include many bullets in full — real names, numbers, dates
                # must survive the trip to the LLM.
                for b in sub.get("bullets", [])[:10]:
                    lines.append(f"    - {b[:300]}")
                # Include substantial paragraph content (not 300-char snippets)
                for p in sub.get("paragraphs", [])[:3]:
                    lines.append(f"    {p[:600]}")
                if sub.get("key_terms"):
                    lines.append(f"    Key terms: {', '.join(sub['key_terms'][:10])}")
                # Include ALL numerical data blocks — these drive chart decisions
                if sub.get("has_numerical_data") and sub.get("numerical_data"):
                    for nd in sub["numerical_data"][:4]:
                        vals = nd.get("values", {})
                        if vals:
                            val_str = ", ".join(f"{k}:{v}" for k, v in list(vals.items())[:10])
                            lines.append(f"    [NUMERIC DATA: {nd.get('context','')} — {val_str}]")
                if sub.get("tables"):
                    tbl = sub["tables"][0]
                    hdrs = " | ".join(tbl.get("headers", [])[:8])
                    first_rows = " / ".join(
                        " | ".join(str(c) for c in row[:8])
                        for row in tbl.get("rows", [])[:5]
                    )
                    lines.append(f"    [TABLE headers: {hdrs}]")
                    if first_rows:
                        lines.append(f"    [TABLE rows: {first_rows}]")

        doc_summary = "\n".join(lines)

        return f"""You are a presentation content analyst. Extract the most important, SPECIFIC content from this document.

DOCUMENT:
{doc_summary}

OUTPUT a JSON object with this exact schema (no markdown, no explanation):
{{
  "title": "document title (under 10 words)",
  "subtitle": "one sentence, max 15 words",
  "executive_summary_bullets": ["specific insight with real fact/number", ...],
  "key_sections": [
    {{
      "heading": "section heading verbatim from document",
      "key_insights": ["specific insight with name/number/fact", ...],
      "visual_type": "chart|table|timeline|process_flow|comparison|none",
      "chart_data": {{
        "chart_type": "bar|line|pie|area",
        "categories": ["cat1", "cat2"],
        "series": [{{"name": "label", "values": [10, 20]}}]
      }},
      "table_data": {{"headers": ["Col1"], "rows": [["val"]]}},
      "process_steps": [{{"number": "01", "heading": "step name", "description": "brief detail"}}],
      "comparison": {{
        "left_heading": "Left Side",
        "left_points": ["specific point"],
        "right_heading": "Right Side",
        "right_points": ["specific point"]
      }}
    }}
  ],
  "global_stats": [{{"value": "6.8x", "label": "descriptive label"}}],
  "suggested_slide_count": 12
}}

CRITICAL QUALITY RULES — violations will make the output useless:
1. key_insights MUST contain specific facts: company names, numbers, percentages, dates, product names, dollar amounts.
   BAD: "Accenture expanded its capabilities"
   GOOD: "NeuraFlash acquisition (2025) added Salesforce agentic AI integration across global markets"
2. Each section MUST have 5 to 7 key_insights. NEVER fewer than 5. These are the bullets that will appear on that section's slide.
3. Each insight MUST be a COMPLETE SENTENCE of 15 to 25 words — not a single word, not a phrase, not a key term.
   BAD: "Halfspace"
   BAD: "AI cybersecurity"
   GOOD: "Halfspace acquisition in March 2025 expanded Accenture's AI footprint into the Nordics via a Denmark-based center"
4. executive_summary_bullets MUST be 4 to 6 specific findings, each 15-25 words with a real number or company name. They MUST NOT restate the document subtitle.
5. global_stats: extract 4 to 6 actual numbers found in the document (e.g. "$6.6B", "326", "$5.9B", "7%"). Each needs a descriptive label.
6. visual_type priority:
   - "chart" if [NUMERIC DATA] block exists with 2+ data points showing a trend or comparison
   - "table" if [TABLE headers] block exists and data does NOT suit a chart
   - "timeline" if content describes sequential phases, years, or numbered steps over time
   - "process_flow" if content describes a workflow or step-by-step process
   - "comparison" if content explicitly contrasts two things (pros/cons, before/after, A vs B)
   - "none" otherwise
7. chart_data: include ONLY when visual_type is "chart", use real numbers from [NUMERIC DATA] blocks
8. table_data: include ONLY when visual_type is "table", copy real headers/rows from [TABLE] blocks verbatim
9. process_steps: include ONLY when visual_type is "timeline" or "process_flow" — 4-6 steps with concrete descriptions
10. comparison: include ONLY when visual_type is "comparison" — 3+ specific points per side
11. suggested_slide_count: 10-15 based on number of major sections
12. Different sections MUST produce different insights. Do not repeat the same sentence across sections.
"""

    # ------------------------------------------------------------------
    # Post-processing
    # ------------------------------------------------------------------

    def _repair(self, result: dict, parsed: dict) -> None:
        """Fix common LLM omissions in-place.

        Args:
            result: Parsed output from the LLM (mutated in-place).
            parsed: Original parsed dict (used as fallback source).
        """
        if not result.get("title"):
            result["title"] = parsed.get("title", "Presentation")

        if not result.get("subtitle"):
            result["subtitle"] = parsed.get("subtitle", "")[:100]

        if not result.get("executive_summary_bullets"):
            exec_sum = parsed.get("executive_summary", "")
            if exec_sum:
                result["executive_summary_bullets"] = [
                    s.strip() for s in exec_sum.split(".") if s.strip()
                ][:4]

        if not isinstance(result.get("key_sections"), list):
            result["key_sections"] = []

        if not isinstance(result.get("global_stats"), list):
            result["global_stats"] = []

        if not isinstance(result.get("suggested_slide_count"), int):
            result["suggested_slide_count"] = 12

        # Clamp slide count
        result["suggested_slide_count"] = max(
            10, min(15, result["suggested_slide_count"])
        )

        # Ensure every key_section has key_insights
        for sec in result.get("key_sections", []):
            if not sec.get("key_insights"):
                sec["key_insights"] = [sec.get("heading", "Key insight")]

    # ------------------------------------------------------------------
    # Rule-based fallback
    # ------------------------------------------------------------------

    def _fallback_extract(self, parsed: dict) -> dict:
        """Build extracted_content without the LLM.

        Maps the parsed dict directly to the extracted_content schema using
        simple heuristics. Called when the LLM is unavailable or fails.

        The goal here is to preserve as MUCH concrete detail as possible
        (full bullets, real paragraph sentences) rather than produce a
        summary — a lossy fallback is what makes slides look shallow.

        Args:
            parsed: Structured dict from MarkdownParser.

        Returns:
            extracted_content dict.
        """
        import re

        logger.info("ContentExtractor: building rule-based extraction.")

        def _clean(text: str) -> str:
            """Strip markdown footnote refs like [1](url) and collapse whitespace."""
            t = re.sub(r"\[\d+\]\([^)]*\)", "", text)
            t = re.sub(r"\s+", " ", t).strip().rstrip(",;:")
            return t

        def _sentences(text: str) -> list[str]:
            """Split a paragraph into sentences, keeping specific/factual ones."""
            parts = re.split(r"(?<=[.!?])\s+", text.strip())
            return [s for s in (p.strip() for p in parts) if len(s.split()) >= 6]

        # -------- Executive summary bullets (1-2 sentences each, up to 6) --
        exec_sum = parsed.get("executive_summary", "")
        exec_bullets: list[str] = []
        if exec_sum:
            for sent in _sentences(_clean(exec_sum)):
                exec_bullets.append(sent)
                if len(exec_bullets) >= 6:
                    break

        # -------- Per-section insight collection ---------------------------
        key_sections: list[dict] = []
        global_stats: list[dict] = []
        seen_insights: set[str] = set()

        _SKIP = ("table of contents", "contents", "references", "bibliography")

        for sec in parsed.get("sections", []):
            heading_low = sec.get("heading", "").lower().lstrip("0123456789. ").strip()
            if any(heading_low.startswith(s) for s in _SKIP):
                continue

            insights: list[str] = []
            visual_type = "none"
            chart_data = None
            table_data = None
            process_steps: list[dict] = []
            all_paragraphs: list[str] = []

            for sub in sec.get("subsections", []):
                # Keep FULL bullet text — real facts live here.
                for b in sub.get("bullets", []):
                    cleaned = _clean(b)
                    if len(cleaned.split()) >= 4 and cleaned not in seen_insights:
                        insights.append(cleaned)
                        seen_insights.add(cleaned)
                        if len(insights) >= 7:
                            break
                # Mine paragraph sentences for more insights.
                for p in sub.get("paragraphs", []):
                    all_paragraphs.append(_clean(p))
                    if len(insights) >= 7:
                        continue
                    for sent in _sentences(_clean(p)):
                        if sent in seen_insights:
                            continue
                        # Prefer sentences that contain a number or proper noun
                        if re.search(r"\d|[A-Z][a-z]+[A-Z]|[A-Z]{2,}", sent):
                            insights.append(sent)
                            seen_insights.add(sent)
                            if len(insights) >= 7:
                                break

                # Chart opportunity
                if sub.get("has_numerical_data") and sub.get("numerical_data") and visual_type == "none":
                    nd = sub["numerical_data"][0]
                    vals = nd.get("values", {})
                    if len(vals) >= 2:
                        visual_type = "chart"
                        chart_data = {
                            "chart_type": "bar",
                            "categories": [str(k) for k in vals.keys()],
                            "series": [{
                                "name": nd.get("context", "Values"),
                                "values": [float(v) for v in vals.values()],
                            }],
                        }
                        for k, v in list(vals.items())[:3]:
                            label = f"{nd.get('context', '')} ({k})".strip()
                            global_stats.append({"value": str(v), "label": label or f"Year {k}"})

                # Table opportunity
                if sub.get("tables") and visual_type == "none":
                    visual_type = "table"
                    table_data = sub["tables"][0]

            # Always produce AT LEAST 5 insights so downstream slides are rich.
            if len(insights) < 5:
                # Fill from section intro text
                for sent in _sentences(_clean(sec.get("content", ""))):
                    if sent not in seen_insights and len(insights) < 5:
                        insights.append(sent)
                        seen_insights.add(sent)
            if len(insights) < 5 and all_paragraphs:
                # Last resort: first sentences of remaining paragraphs
                for p in all_paragraphs:
                    for sent in _sentences(p):
                        if sent not in seen_insights and len(insights) < 5:
                            insights.append(sent)
                            seen_insights.add(sent)
                            break
            if not insights:
                insights = [f"{sec['heading']} — key topic in the document"]

            section_entry: dict = {
                "heading": sec["heading"],
                "key_insights": insights[:7],
                "visual_type": visual_type,
            }
            if chart_data:
                section_entry["chart_data"] = chart_data
            if table_data:
                section_entry["table_data"] = table_data
            if process_steps:
                section_entry["process_steps"] = process_steps

            key_sections.append(section_entry)

        # -------- Global stats: mine standalone dollar / percent figures ---
        if len(global_stats) < 4 and exec_sum:
            stat_pat = re.compile(
                r"(\$\s?[\d,.]+\s?(?:billion|million|B|M|trillion|T|k)?|"
                r"\d+(?:\.\d+)?\s?(?:%|percent)|"
                r"\d{2,4}\+?\s?(?:acquisitions|professionals|companies|deals))",
                re.IGNORECASE,
            )
            seen_values: set[str] = {s.get("value", "") for s in global_stats}
            for m in stat_pat.finditer(exec_sum):
                val = m.group(1).strip()
                if val in seen_values:
                    continue
                seen_values.add(val)
                # Take the surrounding clause and use the words AFTER the
                # match as the label (e.g. "$6.6 billion invested in FY24"
                # → label "invested in FY24").
                end_ctx = exec_sum[m.end() : m.end() + 60]
                label_words = re.split(r"[,;.]", end_ctx)[0].split()[:6]
                label = " ".join(label_words).strip().rstrip(":")
                if not label:
                    start = max(0, m.start() - 40)
                    label = exec_sum[start : m.start()].split()[-5:]
                    label = " ".join(label).strip()
                global_stats.append({"value": val, "label": label[:50] or "metric"})
                if len(global_stats) >= 6:
                    break
        # Final dedup by value+label so identical stats never repeat on slide
        _seen: set[tuple[str, str]] = set()
        deduped: list[dict] = []
        for s in global_stats:
            key = (s.get("value", ""), s.get("label", ""))
            if key in _seen:
                continue
            _seen.add(key)
            deduped.append(s)
        global_stats = deduped

        total_sections = len(key_sections)
        suggested = max(10, min(15, total_sections + 4))

        return {
            "title": parsed.get("title", "Presentation"),
            "subtitle": _clean(parsed.get("subtitle", ""))[:200],
            "executive_summary_bullets": exec_bullets or [
                _clean(sec.get("content", "") or sec.get("heading", ""))
                for sec in parsed.get("sections", [])[:4]
            ],
            "key_sections": key_sections,
            "global_stats": global_stats[:6],
            "suggested_slide_count": suggested,
        }
