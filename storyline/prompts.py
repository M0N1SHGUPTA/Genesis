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
  "cover"             — Title + subtitle only (always slide 1)
  "agenda"            — Table of contents
  "executive_summary" — Key findings in two columns
  "section_divider"   — Visual break between major sections
  "content"           — Main content slide
  "chart"             — Data visualisation slide
  "table"             — Table data slide
  "conclusion"        — Key takeaways (second-to-last)
  "thank_you"         — Closing slide (always last)

LAYOUT OPTIONS (only for "content" and "executive_summary" slides):
  "two_column"    — Left column + right column
  "three_cards"   — Three equal cards in a row
  "key_stats"     — 2–4 big numbers with labels
  "timeline"      — Numbered steps on a line
  "process_flow"  — Boxes with arrows
  "comparison"    — Two contrasting colored columns
  "icon_list"     — Numbered circles + text rows
  "single_focus"  — One key message + bullets

CHART TYPES (only for "chart" slides):
  "bar" / "pie" / "line" / "area"
"""


# ---------------------------------------------------------------------------
# Main blueprint prompt
# ---------------------------------------------------------------------------

def build_blueprint_prompt(parsed: dict, target_slides: int | None) -> str:
    """Build the full prompt for generating a slide blueprint.

    Args:
        parsed: Structured dict from Stage 1 (MarkdownParser).
        target_slides: Requested slide count, or None to let LLM decide.

    Returns:
        Complete prompt string.
    """
    slide_hint = (
        f"The presentation MUST have exactly {target_slides} slides."
        if target_slides is not None
        else "Choose a slide count between 10 and 15 based on content density."
    )

    parsed_json = json.dumps(parsed, indent=2, ensure_ascii=False)

    return f"""You are a professional presentation designer. Convert this document into a slide blueprint JSON.

## DOCUMENT CONTENT
{parsed_json}

## SLIDE COUNT RULE
{slide_hint}
Total slides must be between 10 and 15.

## MANDATORY SLIDE ORDER
1. Slide 1: type "cover" — ONLY title + subtitle (max 20 words in subtitle)
2. Slide 2: type "executive_summary" — extracted from executive_summary field
3. Middle slides: section_divider, content, chart, table (cover all major sections)
4. Second-to-last: type "conclusion"
5. Last slide: type "thank_you"

## STRICT CONTENT RULES

### Cover slide:
- "title": the document title (under 10 words)
- "subtitle": ONE sentence only, MAX 20 WORDS summarising the document. NO executive summary text here.

### Executive summary slide:
- layout: "two_column"
- "left": {{ "heading": "Key Findings", "points": [3-4 bullets FROM the executive_summary field] }}
- "right": {{ "heading": "Implications", "points": [3-4 bullets FROM the sections content] }}
- points arrays MUST be non-empty. Extract real content from the document.

### All content slides:
- MAX 6 bullet points per slide or per column
- MAX 15 WORDS per bullet point — summarise aggressively
- NEVER leave points/cards/items arrays empty

### Layout variety (MANDATORY):
- NEVER use the same layout on two consecutive slides
- Rotate through: two_column → three_cards → key_stats → timeline → icon_list → single_focus
- Agenda slide: list ALL section headings as bullets under "points"
- Use "key_stats" for any slide with numerical data (percentages, dollar amounts, counts)
- Use "three_cards" for slides with exactly 3 concepts or categories
- Use "two_column" for comparison or before/after content

### VISUAL LAYOUT PREFERENCE (CRITICAL):
- PREFER infographic/visual layouts (icon_list, three_cards, key_stats, timeline, process_flow) over text-heavy layouts (single_focus, two_column)
- Only use single_focus or two_column when content truly does not fit any visual pattern
- When a section lists 3–5 items with descriptions, use icon_list or three_cards — NOT bullet points in single_focus
- When content describes a sequence, phases, or steps, use timeline or process_flow
- When data has success/failure, pros/cons, or two contrasting sides, use comparison layout
- Slides should feel like DESIGNED slides, not documents — every slide needs visual structure (cards, icons, charts), not just text

### Charts (IMPORTANT — prefer charts over number-bullets):
- When content lists numerical outcomes (revenue, headcount, percentages, growth rates), ALWAYS create a chart slide rather than putting numbers in bullet points
- When comparing numerical success vs failure counts, use a bar chart
- Only create "chart" slides when the parsed content has actual numerical_data
- "categories" MUST be a list of strings: ["2020", "2021", "2022"]
- "series" MUST be: [{{"name": "label", "values": [10, 25, 47]}}]
- values MUST be plain numbers, not strings

{SLIDE_TYPE_REFERENCE}

## OUTPUT FORMAT
Return ONLY valid JSON. No markdown fences, no explanation text.

{{
  "presentation_title": "string",
  "total_slides": 12,
  "slides": [
    {{
      "slide_number": 1,
      "type": "cover",
      "title": "Short Title Here",
      "subtitle": "One sentence max 20 words describing the document."
    }},
    {{
      "slide_number": 2,
      "type": "executive_summary",
      "layout": "two_column",
      "title": "Executive Summary",
      "left": {{
        "heading": "Key Findings",
        "points": ["Finding one from the document", "Finding two from the document", "Finding three"]
      }},
      "right": {{
        "heading": "Implications",
        "points": ["Implication one", "Implication two", "Implication three"]
      }}
    }},
    {{
      "slide_number": 3,
      "type": "agenda",
      "title": "Agenda",
      "points": ["Section 1 heading", "Section 2 heading", "Section 3 heading", "Section 4 heading"]
    }},
    {{
      "slide_number": 4,
      "type": "section_divider",
      "title": "Section Name",
      "subtitle": "Brief description"
    }},
    {{
      "slide_number": 5,
      "type": "content",
      "layout": "three_cards",
      "title": "Slide Title",
      "cards": [
        {{"number": "01", "heading": "Card One", "points": ["Point A", "Point B"]}},
        {{"number": "02", "heading": "Card Two", "points": ["Point A", "Point B"]}},
        {{"number": "03", "heading": "Card Three", "points": ["Point A", "Point B"]}}
      ]
    }},
    {{
      "slide_number": 6,
      "type": "content",
      "layout": "two_column",
      "title": "Slide Title",
      "left": {{"heading": "Left Heading", "points": ["Point 1", "Point 2", "Point 3"]}},
      "right": {{"heading": "Right Heading", "points": ["Point 1", "Point 2", "Point 3"]}}
    }},
    {{
      "slide_number": 7,
      "type": "content",
      "layout": "key_stats",
      "title": "Key Statistics",
      "stats": [
        {{"value": "47%", "label": "Market share"}},
        {{"value": "$2.1B", "label": "Total investment"}},
        {{"value": "3.4x", "label": "Growth rate"}}
      ]
    }},
    {{
      "slide_number": 8,
      "type": "content",
      "layout": "icon_list",
      "title": "Key Points",
      "items": [
        {{"number": "01", "heading": "Item One", "description": "Brief description here"}},
        {{"number": "02", "heading": "Item Two", "description": "Brief description here"}},
        {{"number": "03", "heading": "Item Three", "description": "Brief description here"}}
      ]
    }},
    {{
      "slide_number": 9,
      "type": "chart",
      "title": "Chart Title",
      "chart_type": "bar",
      "data": {{
        "categories": ["2020", "2021", "2022", "2023"],
        "series": [{{"name": "Value ($B)", "values": [10, 25, 47, 68]}}]
      }},
      "caption": "Short insight about this data"
    }},
    {{
      "slide_number": 10,
      "type": "content",
      "layout": "timeline",
      "title": "Timeline",
      "steps": [
        {{"number": "01", "heading": "Step One", "description": "Brief"}},
        {{"number": "02", "heading": "Step Two", "description": "Brief"}},
        {{"number": "03", "heading": "Step Three", "description": "Brief"}}
      ]
    }},
    {{
      "slide_number": 11,
      "type": "conclusion",
      "layout": "single_focus",
      "title": "Key Takeaways",
      "focus": "One sentence summarising the main conclusion",
      "points": ["Takeaway 1", "Takeaway 2", "Takeaway 3", "Takeaway 4"]
    }},
    {{
      "slide_number": 12,
      "type": "thank_you",
      "title": "Thank You",
      "subtitle": "Contact or closing remark"
    }}
  ]
}}

REMEMBER:
- Cover subtitle = MAX 20 WORDS
- Every points/cards/items/stats array MUST have real content — never empty
- No two consecutive slides with the same layout
- Chart values must be numbers, not strings
"""


# ---------------------------------------------------------------------------
# Summary prompt (for large documents)
# ---------------------------------------------------------------------------

def build_summary_prompt(parsed: dict, target_slides: int | None) -> str:
    """Condensed prompt — rich enough for substantive slide content.

    Sends ~7k input tokens, leaving 5k budget for the 3.5k output cap.
    Includes: full exec summary, 4 bullets per subsection, key terms,
    numerical data values, paragraph excerpts.

    Args:
        parsed: Full structured dict from Stage 1.
        target_slides: Requested slide count, or None.

    Returns:
        Condensed prompt string.
    """
    slide_hint = (
        f"The presentation MUST have exactly {target_slides} slides."
        if target_slides is not None
        else "Choose a slide count between 10 and 15."
    )

    lines: list[str] = []
    lines.append(f"Title: {parsed.get('title', 'Untitled')}")
    lines.append(f"Subtitle: {parsed.get('subtitle', '')[:200]}")

    # Full executive summary (up to 1200 chars — most important context)
    exec_sum = parsed.get("executive_summary", "")
    if exec_sum:
        lines.append(f"\nExecutive Summary:\n{exec_sum[:1200]}")

    for sec in parsed.get("sections", []):
        lines.append(f"\n## {sec['heading']}")
        # Include section-level content preview
        if sec.get("content"):
            lines.append(f"  {sec['content'][:200]}")

        for sub in sec.get("subsections", []):
            if sub.get("heading"):
                lines.append(f"  ### {sub['heading']}")
            # Up to 4 bullets per subsection
            for bullet in sub.get("bullets", [])[:4]:
                lines.append(f"    - {bullet}")
            # Up to 2 paragraph excerpts
            for para in sub.get("paragraphs", [])[:2]:
                lines.append(f"    {para[:200]}")
            # Key terms (bold/italic words)
            if sub.get("key_terms"):
                terms = ", ".join(sub["key_terms"][:6])
                lines.append(f"    Key terms: {terms}")
            # Numerical data — include actual values, not just a flag
            if sub.get("has_numerical_data") and sub.get("numerical_data"):
                for nd in sub["numerical_data"][:2]:
                    vals = nd.get("values", {})
                    if vals:
                        val_str = ", ".join(f"{k}:{v}" for k, v in list(vals.items())[:6])
                        lines.append(f"    [DATA: {nd.get('context','')} — {val_str}] → chart slide")
            if sub.get("tables"):
                tbl = sub["tables"][0]
                hdrs = " | ".join(tbl.get("headers", []))
                lines.append(f"    [TABLE: {hdrs}] → table slide")

    condensed = "\n".join(lines)

    return f"""You are a professional presentation designer. Convert this document into a detailed slide blueprint JSON.

## DOCUMENT CONTENT
{condensed}

## SLIDE COUNT
{slide_hint}

## MANDATORY STRUCTURE
- Slide 1: cover — title + ONE sentence subtitle (max 20 words, no exec summary text)
- Slide 2: executive_summary, layout two_column — extract REAL insights from Executive Summary above
- Slide 3: agenda — "points" array listing ALL section headings
- Second-to-last: conclusion with key takeaways
- Last: thank_you

## CONTENT RULES
- Every bullet max 15 words — summarise aggressively but keep specific facts and numbers
- Include actual numbers/percentages/values from the document in bullets where possible
- MAX 6 bullets per column or card
- NEVER leave points/cards/items arrays empty — always use real content from the document
- NEVER use same layout on two consecutive slides

## LAYOUT ROTATION
Rotate through these layouts across content slides:
two_column → three_cards → key_stats → timeline → icon_list → single_focus

Use key_stats when slide has specific numbers/percentages.
Use three_cards when slide has exactly 3 concepts.
Use two_column for comparisons or before/after.

## VISUAL LAYOUT PREFERENCE (CRITICAL)
PREFER infographic layouts (icon_list, three_cards, key_stats, timeline, process_flow) over text-heavy ones (single_focus, two_column).
When 3-5 items have descriptions → use icon_list or three_cards.
When content describes steps/phases → use timeline or process_flow.
When data has pros/cons or contrasts → use comparison.
Slides must feel DESIGNED, not like documents.

## CHART RULES
- When content has numerical outcomes (revenue, headcount, growth), ALWAYS create a chart slide instead of number-bullets
- Only create chart slides when DATA entries appear above
- categories: list of strings, values: list of numbers (not strings)

## OUTPUT
Return ONLY valid JSON. No markdown fences, no explanation text.
Schema: same as standard blueprint — presentation_title, total_slides, slides array.
Each slide needs: slide_number, type, title, and layout-specific fields.
"""


# ---------------------------------------------------------------------------
# Correction prompt (retry on invalid JSON)
# ---------------------------------------------------------------------------

def build_correction_prompt(bad_response: str, error: str, target_slides: int | None) -> str:
    """Ask the LLM to fix its previous invalid JSON response.

    Args:
        bad_response: The raw invalid response from the previous attempt.
        error: The JSON parse error message.
        target_slides: Required slide count constraint.

    Returns:
        Correction prompt string.
    """
    slide_constraint = (
        f"total_slides must be exactly {target_slides}."
        if target_slides is not None
        else "total_slides must be between 10 and 15."
    )

    return f"""Your previous response had invalid JSON.

Error: {error}

Previous response (first 2000 chars):
{bad_response[:2000]}

Fix it and return ONLY valid JSON. No markdown fences, no explanation.
Rules:
- {slide_constraint}
- Cover subtitle MAX 20 words
- All points/cards/items arrays must be non-empty
- Chart values must be numbers not strings
- Close all brackets and quote all strings
"""
