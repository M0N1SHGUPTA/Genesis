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

# Token budget: output is compact (key insights only), so 1500 is enough.
_MAX_TOKENS = 1500


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

        Sends a compact but content-rich view of the parsed document so the LLM
        has enough real facts (company names, numbers, dates) to produce
        specific, non-generic insights.
        """
        lines: list[str] = [
            f"Title: {parsed.get('title', 'Untitled')}",
            f"Subtitle: {parsed.get('subtitle', '')[:200]}",
        ]

        exec_sum = parsed.get("executive_summary", "")
        if exec_sum:
            lines.append(f"\nExecutive Summary:\n{exec_sum[:1200]}")

        for sec in parsed.get("sections", []):
            lines.append(f"\n## {sec['heading']}")
            if sec.get("content"):
                lines.append(f"  {sec['content'][:300]}")

            for sub in sec.get("subsections", []):
                if sub.get("heading"):
                    lines.append(f"  ### {sub['heading']}")
                # Include more bullets so real names/numbers survive condensation
                for b in sub.get("bullets", [])[:6]:
                    lines.append(f"    - {b}")
                # Include more paragraph content
                for p in sub.get("paragraphs", [])[:2]:
                    lines.append(f"    {p[:300]}")
                if sub.get("key_terms"):
                    lines.append(f"    Key terms: {', '.join(sub['key_terms'][:8])}")
                # Include ALL numerical data blocks — these drive chart decisions
                if sub.get("has_numerical_data") and sub.get("numerical_data"):
                    for nd in sub["numerical_data"][:3]:
                        vals = nd.get("values", {})
                        if vals:
                            val_str = ", ".join(f"{k}:{v}" for k, v in list(vals.items())[:8])
                            lines.append(f"    [NUMERIC DATA: {nd.get('context','')} — {val_str}]")
                if sub.get("tables"):
                    tbl = sub["tables"][0]
                    hdrs = " | ".join(tbl.get("headers", [])[:6])
                    first_rows = " / ".join(
                        " | ".join(str(c) for c in row[:6])
                        for row in tbl.get("rows", [])[:3]
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
1. key_insights MUST contain specific facts: company names, numbers, percentages, dates, product names
   BAD: "Accenture expanded its capabilities"
   GOOD: "NeuraFlash acquisition added Salesforce AI expertise in 2023"
2. executive_summary_bullets MUST be specific findings, not document titles or subtitles
3. global_stats: extract actual numbers found in the document (e.g. "$69.7B revenue", "80 acquisitions")
4. visual_type priority:
   - "chart" if [NUMERIC DATA] block exists with 2+ data points showing a trend or comparison
   - "table" if [TABLE headers] block exists and data does NOT suit a chart
   - "timeline" if content describes sequential phases, years, or numbered steps over time
   - "process_flow" if content describes a workflow or step-by-step process
   - "comparison" if content explicitly contrasts two things (pros/cons, before/after, A vs B)
   - "none" otherwise
5. chart_data: include ONLY when visual_type is "chart", use real numbers from [NUMERIC DATA] blocks
6. table_data: include ONLY when visual_type is "table", copy real headers/rows from [TABLE] blocks
7. process_steps: include ONLY when visual_type is "timeline" or "process_flow"
8. comparison: include ONLY when visual_type is "comparison"
9. suggested_slide_count: 10-15 based on number of major sections
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

        Args:
            parsed: Structured dict from MarkdownParser.

        Returns:
            extracted_content dict.
        """
        logger.info("ContentExtractor: building rule-based extraction.")

        exec_sum = parsed.get("executive_summary", "")
        exec_bullets: list[str] = []
        if exec_sum:
            import re
            sentences = re.split(r"(?<=[.!?])\s+", exec_sum.strip())
            for s in sentences:
                words = s.split()
                if words:
                    exec_bullets.append(" ".join(words[:12]))
                if len(exec_bullets) >= 4:
                    break

        key_sections: list[dict] = []
        global_stats: list[dict] = []

        for sec in parsed.get("sections", []):
            # Gather insights from subsection bullets
            insights: list[str] = []
            visual_type = "none"
            chart_data = None
            table_data = None
            process_steps: list[dict] = []

            for sub in sec.get("subsections", []):
                for b in sub.get("bullets", [])[:2]:
                    words = b.split()
                    insights.append(" ".join(words[:12]))
                if not insights and sub.get("paragraphs"):
                    text = sub["paragraphs"][0]
                    words = text.split()
                    insights.append(" ".join(words[:12]))

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
                        # Collect global stats from numerical data
                        for k, v in list(vals.items())[:2]:
                            global_stats.append({"value": str(v), "label": f"{nd.get('context', '')} ({k})"})

                # Table opportunity
                if sub.get("tables") and visual_type == "none":
                    visual_type = "table"
                    table_data = sub["tables"][0]

            if not insights:
                insights = [sec["heading"]]

            section_entry: dict = {
                "heading": sec["heading"],
                "key_insights": insights[:4],
                "visual_type": visual_type,
            }
            if chart_data:
                section_entry["chart_data"] = chart_data
            if table_data:
                section_entry["table_data"] = table_data
            if process_steps:
                section_entry["process_steps"] = process_steps

            key_sections.append(section_entry)

        total_sections = len(key_sections)
        suggested = max(10, min(15, total_sections + 4))  # +4 for cover/exec/conclusion/thanks

        return {
            "title": parsed.get("title", "Presentation"),
            "subtitle": (parsed.get("subtitle", "")[:100]),
            "executive_summary_bullets": exec_bullets or ["Key findings from the document"],
            "key_sections": key_sections,
            "global_stats": global_stats[:4],
            "suggested_slide_count": suggested,
        }
