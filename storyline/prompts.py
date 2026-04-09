"""
prompts.py — LLM prompt templates for the storyline generator.

All prompt-building logic lives here so generator.py stays clean.
"""

from __future__ import annotations

import json


# ---------------------------------------------------------------------------
# Slide type & layout reference (shown to the LLM)
# ---------------------------------------------------------------------------

SLIDE_TYPE_REFERENCE = """
SLIDE TYPES (use exactly these strings):
  "cover"             — Title + subtitle (always slide 1)
  "agenda"            — Table of contents (optional, use only if markdown has one)
  "executive_summary" — Key findings summary (use if executive_summary field is non-empty)
  "section_divider"   — Visual break between major sections
  "content"           — Main content slide (text, bullets, infographics)
  "chart"             — Data visualisation slide
  "table"             — Table data slide
  "conclusion"        — Key takeaways (second-to-last)
  "thank_you"         — Closing slide (always last)

LAYOUT OPTIONS (only for "content" and "executive_summary" type slides):
  "two_column"    — Left column + right column (good for comparisons, summaries)
  "three_cards"   — Three equal cards in a row (good for 3 concepts / pillars)
  "key_stats"     — 2–4 big numbers with labels (good for statistics)
  "timeline"      — Numbered steps on a horizontal line (good for sequences)
  "process_flow"  — Boxes with arrows (good for workflows)
  "comparison"    — Side-by-side comparison blocks (good for pros/cons)
  "icon_list"     — Numbered circles + heading + description rows (3–4 items)
  "single_focus"  — One key message + supporting bullets

CHART TYPES (only for "chart" type slides):
  "bar"   — Comparing categories
  "pie"   — Part-of-whole / proportions
  "line"  — Trends over time
  "area"  — Cumulative trends over time
"""


# ---------------------------------------------------------------------------
# Main blueprint prompt
# ---------------------------------------------------------------------------

def build_blueprint_prompt(parsed: dict, target_slides: int | None) -> str:
    """Build the full prompt for generating a slide blueprint from parsed markdown.

    Args:
        parsed: The structured dict produced by Stage 1 (MarkdownParser).
        target_slides: Requested slide count, or None to let the LLM decide.

    Returns:
        A complete prompt string ready to send to the LLM.
    """
    slide_hint = (
        f"The presentation MUST have exactly {target_slides} slides."
        if target_slides is not None
        else "Choose a slide count between 10 and 15 based on content density."
    )

    parsed_json = json.dumps(parsed, indent=2, ensure_ascii=False)

    return f"""You are a professional presentation designer. Your job is to convert structured document content into a slide-by-slide blueprint for a PowerPoint presentation.

## PARSED DOCUMENT CONTENT
{parsed_json}

## YOUR TASK
Convert the above content into a complete slide blueprint JSON. Follow ALL rules below exactly.

## SLIDE COUNT
{slide_hint}
The count must be between 10 and 15 (inclusive).

## MANDATORY SLIDE ORDER
1. Slide 1: type "cover" — always first
2. Slide 2: type "agenda" — only if the document has a clear table of contents
3. Next: type "executive_summary" — only if executive_summary field is non-empty
4. Middle slides: mix of "section_divider", "content", "chart", "table"
5. Second-to-last: type "conclusion"
6. Last slide: type "thank_you" — always last

## CONTENT RULES
- MAX 6 bullet points per slide or column
- MAX 15 words per bullet point
- NO walls of text — summarise aggressively
- Every section from the markdown must appear somewhere
- Vary layouts — never use the same layout on two consecutive slides
- When numerical_data exists in the parsed content, create a "chart" slide for it
- When a table exists in the parsed content, create a "table" slide for it

{SLIDE_TYPE_REFERENCE}

## OUTPUT FORMAT
Return ONLY valid JSON. No markdown fences, no explanation, no extra text.
The JSON must exactly match this schema:

{{
  "presentation_title": "string",
  "total_slides": <integer 10-15>,
  "slides": [
    {{
      "slide_number": 1,
      "type": "cover",
      "title": "string",
      "subtitle": "string"
    }},
    {{
      "slide_number": 2,
      "type": "executive_summary",
      "layout": "two_column",
      "title": "string",
      "left": {{
        "heading": "string",
        "points": ["string", "string"]
      }},
      "right": {{
        "heading": "string",
        "points": ["string", "string"]
      }}
    }},
    {{
      "slide_number": 3,
      "type": "section_divider",
      "title": "string",
      "subtitle": "string"
    }},
    {{
      "slide_number": 4,
      "type": "content",
      "layout": "three_cards",
      "title": "string",
      "cards": [
        {{"number": "01", "heading": "string", "points": ["string"]}},
        {{"number": "02", "heading": "string", "points": ["string"]}},
        {{"number": "03", "heading": "string", "points": ["string"]}}
      ]
    }},
    {{
      "slide_number": 5,
      "type": "content",
      "layout": "two_column",
      "title": "string",
      "left": {{"heading": "string", "points": ["string"]}},
      "right": {{"heading": "string", "points": ["string"]}}
    }},
    {{
      "slide_number": 6,
      "type": "content",
      "layout": "key_stats",
      "title": "string",
      "stats": [
        {{"value": "string", "label": "string"}},
        {{"value": "string", "label": "string"}}
      ]
    }},
    {{
      "slide_number": 7,
      "type": "content",
      "layout": "timeline",
      "title": "string",
      "steps": [
        {{"number": "01", "heading": "string", "description": "string"}},
        {{"number": "02", "heading": "string", "description": "string"}}
      ]
    }},
    {{
      "slide_number": 8,
      "type": "content",
      "layout": "icon_list",
      "title": "string",
      "items": [
        {{"number": "01", "heading": "string", "description": "string"}},
        {{"number": "02", "heading": "string", "description": "string"}}
      ]
    }},
    {{
      "slide_number": 9,
      "type": "content",
      "layout": "single_focus",
      "title": "string",
      "focus": "string",
      "points": ["string", "string"]
    }},
    {{
      "slide_number": 10,
      "type": "chart",
      "title": "string",
      "chart_type": "bar",
      "data": {{
        "categories": ["string"],
        "series": [
          {{"name": "string", "values": [0]}}
        ]
      }},
      "caption": "string"
    }},
    {{
      "slide_number": 11,
      "type": "table",
      "title": "string",
      "table": {{
        "headers": ["string"],
        "rows": [["string"]]
      }},
      "caption": "string"
    }},
    {{
      "slide_number": 12,
      "type": "conclusion",
      "layout": "single_focus",
      "title": "Key Takeaways",
      "focus": "string",
      "points": ["string", "string", "string"]
    }},
    {{
      "slide_number": 13,
      "type": "thank_you",
      "title": "Thank You",
      "subtitle": "string"
    }}
  ]
}}

Use only the slide types and layouts that fit the actual content. The schema above shows all possible shapes — only include the fields relevant to each slide's type and layout.
"""


# ---------------------------------------------------------------------------
# Summary prompt (for large documents that exceed context window)
# ---------------------------------------------------------------------------

def build_summary_prompt(parsed: dict, target_slides: int | None) -> str:
    """Build a condensed prompt using only headings + first 2 bullets per section.

    Used as a fallback when the full parsed JSON is too large for the LLM.

    Args:
        parsed: The full structured dict from Stage 1.
        target_slides: Requested slide count, or None to let the LLM decide.

    Returns:
        A condensed prompt string.
    """
    slide_hint = (
        f"The presentation MUST have exactly {target_slides} slides."
        if target_slides is not None
        else "Choose a slide count between 10 and 15 based on content density."
    )

    # Build a condensed summary of the document
    lines: list[str] = []
    lines.append(f"Title: {parsed.get('title', 'Untitled')}")
    lines.append(f"Subtitle: {parsed.get('subtitle', '')}")

    exec_sum = parsed.get("executive_summary", "")
    if exec_sum:
        lines.append(f"Executive Summary: {exec_sum[:300]}")

    for sec in parsed.get("sections", []):
        lines.append(f"\n## {sec['heading']}")
        for sub in sec.get("subsections", []):
            if sub["heading"]:
                lines.append(f"  ### {sub['heading']}")
            for bullet in sub.get("bullets", [])[:2]:
                lines.append(f"    - {bullet}")
            for para in sub.get("paragraphs", [])[:1]:
                lines.append(f"    {para[:150]}")
            if sub.get("has_numerical_data"):
                lines.append("    [NUMERICAL DATA AVAILABLE — suggest a chart slide]")
            if sub.get("tables"):
                lines.append("    [TABLE AVAILABLE — suggest a table slide]")

    condensed = "\n".join(lines)

    return f"""You are a professional presentation designer. Convert this document outline into a slide blueprint.

## DOCUMENT OUTLINE (condensed)
{condensed}

## YOUR TASK
Create a complete slide blueprint JSON for this content.

## SLIDE COUNT
{slide_hint}
The count must be between 10 and 15 (inclusive).

{SLIDE_TYPE_REFERENCE}

## RULES
- Follow mandatory slide order: cover → (agenda) → (executive_summary) → content → conclusion → thank_you
- MAX 6 bullet points per slide, MAX 15 words per bullet
- Vary layouts — no two consecutive slides with the same layout
- Create chart slides where numerical data is noted
- Create table slides where tables are noted

Return ONLY valid JSON matching the blueprint schema. No markdown, no explanation.
"""


# ---------------------------------------------------------------------------
# Correction prompt (used on retry when LLM returns invalid JSON)
# ---------------------------------------------------------------------------

def build_correction_prompt(bad_response: str, error: str, target_slides: int | None) -> str:
    """Build a prompt asking the LLM to fix its previous invalid JSON response.

    Args:
        bad_response: The raw invalid response from the previous attempt.
        error: The JSON parse error message.
        target_slides: Required slide count constraint.

    Returns:
        A correction prompt string.
    """
    slide_constraint = (
        f"The total_slides field must be exactly {target_slides}."
        if target_slides is not None
        else "The total_slides field must be between 10 and 15."
    )

    return f"""Your previous response contained invalid JSON.

Error: {error}

Previous response (truncated to 2000 chars):
{bad_response[:2000]}

Fix the JSON and return ONLY valid JSON. No markdown fences, no explanation.
{slide_constraint}
Ensure all brackets and braces are properly closed and all strings are quoted.
"""
