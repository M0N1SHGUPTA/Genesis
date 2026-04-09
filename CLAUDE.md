# CLAUDE.md — MD to PPTX Converter

## Project Overview

Build a CLI tool that converts Markdown (.md) files into professional, visually appealing PowerPoint (.pptx) presentations using a provided Slide Master template. This is for the **Code EZ: Master of Agents** hackathon by EZ.

### Usage
```bash
python main.py --md input.md --template template.pptx --output output.pptx --slides 12
```

### Two Inputs
1. **Markdown file** — the content source (can be up to 5MB, deeply nested headings, tables, numerical data)
2. **Template PPTX** — the Slide Master file that dictates branding, backgrounds, colors, fonts. All output slides MUST inherit from this template.

### One Output
- A valid `.pptx` file (10-15 slides) that opens in PowerPoint, Google Slides, and LibreOffice Impress.

---

## System Architecture

Three-stage pipeline with clear separation of concerns:

```
input.md + template.pptx
        │
        ▼
┌──────────────────────┐
│  Stage 1: PARSER     │  → Structured JSON
│  (parser/md_parser.py)│
└──────────┬───────────┘
           ▼
┌──────────────────────┐
│  Stage 2: STORYLINE  │  → Slide Blueprint JSON
│  (storyline/         │
│   generator.py)      │
└──────────┬───────────┘
           ▼
┌──────────────────────┐
│  Stage 3: RENDERER   │  → output.pptx
│  (renderer/engine.py)│
└──────────────────────┘
```

---

## File Structure

```
md-to-pptx/
├── main.py                  # CLI entry point (argparse)
├── config.py                # Colors, fonts, dimensions, constants
├── parser/
│   ├── __init__.py
│   └── md_parser.py         # Markdown → structured JSON
├── storyline/
│   ├── __init__.py
│   ├── generator.py         # Parsed JSON → slide blueprint via LLM
│   └── prompts.py           # LLM prompt templates
├── renderer/
│   ├── __init__.py
│   ├── engine.py            # Main render loop — loads template, iterates slides
│   ├── layouts.py           # Layout functions (two_column, three_cards, etc.)
│   ├── charts.py            # python-pptx chart generation (bar, pie, line, area)
│   ├── tables.py            # Styled table generation
│   ├── infographics.py      # Process flows, timelines, comparison shapes
│   └── utils.py             # Shared helpers (add_textbox, style_shape, etc.)
├── templates/               # Place Slide Master .pptx files here
├── test_cases/              # Place test .md files here
├── outputs/                 # Generated .pptx files go here
├── requirements.txt
├── .env                     # GROQ_API_KEY=gsk_...
├── README.md
└── CLAUDE.md                # This file
```

---

## Stage 1: Markdown Parser (`parser/md_parser.py`)

### Purpose
Parse the markdown file into a structured JSON representation that the storyline generator can work with.

### Library
Use `mistune` (v3) — it provides an AST renderer out of the box.

```bash
pip install mistune
```

### What to Extract
1. **Title** — first H1 heading
2. **Subtitle** — first paragraph or text immediately after H1
3. **Executive Summary** — content under a heading containing "executive summary" (case-insensitive)
4. **Sections** — each H2 becomes a section
5. **Subsections** — each H3 under an H2 becomes a subsection
6. **Content types per section:**
   - `bullets` — list items
   - `paragraphs` — plain text blocks
   - `tables` — markdown tables with headers and rows
   - `numerical_data` — extracted numbers with context (for chart generation)
   - `key_terms` — bold or emphasized text
7. **Table of Contents** — if present, extract it but don't render it as a slide

### Output Schema
```json
{
  "title": "AI Bubble: Detection, Prevention, and Investment Strategies",
  "subtitle": "A comprehensive guide on global AI bubble risks...",
  "executive_summary": "The report frames the AI bubble through...",
  "sections": [
    {
      "heading": "Introduction to the AI Bubble Concept",
      "level": 2,
      "content": "This section frames the possibility...",
      "subsections": [
        {
          "heading": "Economic bubble theory applied to AI investment dynamics",
          "level": 3,
          "bullets": ["point 1", "point 2"],
          "paragraphs": ["paragraph text..."],
          "tables": [
            {
              "headers": ["Year", "Investment ($B)", "Growth %"],
              "rows": [["2020", "10", "15%"], ["2021", "25", "150%"]]
            }
          ],
          "has_numerical_data": true,
          "numerical_data": [
            {"context": "AI investment", "values": {"2020": 10, "2021": 25, "2022": 47}}
          ]
        }
      ]
    }
  ],
  "total_sections": 15,
  "total_tables": 3,
  "total_numerical_blocks": 5
}
```

### Edge Cases to Handle
- Markdown with no H1 (use filename as title)
- Markdown with no executive summary (skip that slide)
- Deeply nested headings (H4, H5 — flatten into parent H3)
- Tables with irregular columns (pad with empty strings)
- Very large files (5MB) — don't load everything into memory at once
- Missing or malformed table syntax
- Code blocks — skip them (not relevant for presentations)
- Links — extract text only, discard URLs
- Images in markdown — note their alt text but don't try to embed (copyright rules)

---

## Stage 2: Storyline Generator (`storyline/generator.py`)

### Purpose
Use an LLM (Groq API with llama-3.3-70b-versatile) to create an intelligent slide-by-slide blueprint from the parsed content. This is what makes the system "intelligent" rather than a dumb 1:1 transformer.

### Library
```bash
pip install groq
```

### Environment
```
GROQ_API_KEY=gsk_your_key_here
```

### What the LLM Decides
1. **Slide count** (10-15) based on content density
2. **Content per slide** — merges related subsections, summarizes long text
3. **Layout type** for each slide
4. **What becomes a chart** vs text vs table
5. **Content trimming** — max 6 bullet points per slide, each under 15 words
6. **Narrative flow** — coherent storyline, not just section-by-section dump

### LLM Prompt Strategy

The prompt in `prompts.py` should:
1. Include the full parsed JSON (or a summary if it's too large)
2. Specify the exact output JSON schema
3. Enforce slide count range (10-15)
4. Enforce content limits (no walls of text)
5. Request layout type assignments
6. Ask for chart type decisions when numerical data exists

### Slide Types & Layouts the LLM Can Choose From

```
SLIDE TYPES:
- "cover"              → Title + subtitle (uses template Cover layout)
- "agenda"             → Optional table of contents (uses Blank layout)
- "executive_summary"  → Key findings summary (uses Blank layout)
- "section_divider"    → Section break title (uses template Divider layout)
- "content"            → Main content slide (uses Title-only or Blank layout)
- "chart"              → Data visualization slide (uses Blank layout)
- "table"              → Table data slide (uses Blank layout)
- "conclusion"         → Key takeaways (uses Blank layout)
- "thank_you"          → Closing slide (uses template Thank You layout)

LAYOUT OPTIONS (for "content" type slides):
- "two_column"         → Left column + right column
- "three_cards"        → Three equal cards in a row
- "key_stats"          → 2-4 big numbers with labels
- "timeline"           → Numbered steps connected by a line/arrow
- "process_flow"       → Flowchart-style boxes with arrows
- "comparison"         → Side-by-side comparison blocks
- "icon_list"          → Icon circles + text rows (3-4 items)
- "single_focus"       → One key message with supporting points
```

### Output Schema (Slide Blueprint)
```json
{
  "presentation_title": "AI Bubble: Detection & Prevention",
  "total_slides": 12,
  "slides": [
    {
      "slide_number": 1,
      "type": "cover",
      "title": "AI Bubble: Detection, Prevention, and Investment Strategies",
      "subtitle": "A comprehensive guide on global AI bubble risks..."
    },
    {
      "slide_number": 2,
      "type": "executive_summary",
      "layout": "two_column",
      "title": "Executive Summary",
      "left": {
        "heading": "Key Findings",
        "points": ["AI investment shows bubble patterns", "Market concentration rising"]
      },
      "right": {
        "heading": "Recommendations",
        "points": ["Diversify AI portfolios", "Monitor valuation metrics"]
      }
    },
    {
      "slide_number": 5,
      "type": "chart",
      "title": "Global AI Investment Growth",
      "chart_type": "bar",
      "data": {
        "categories": ["2020", "2021", "2022", "2023"],
        "series": [
          {"name": "Investment ($B)", "values": [10, 25, 47, 68]}
        ]
      },
      "caption": "AI investment has grown 6.8x in four years"
    },
    {
      "slide_number": 7,
      "type": "content",
      "layout": "three_cards",
      "title": "Warning Indicators of Overvaluation",
      "cards": [
        {
          "number": "01",
          "heading": "Valuation disconnect",
          "points": ["P/E ratios exceed historical norms", "Revenue growth not matching valuations"]
        },
        {
          "number": "02",
          "heading": "Concentration risk",
          "points": ["Top 5 firms hold 70% of AI market cap", "Single-point-of-failure exposure"]
        },
        {
          "number": "03",
          "heading": "Speculative financing",
          "points": ["Debt-fueled capex cycles", "Circular revenue dependencies"]
        }
      ]
    }
  ]
}
```

### Handling Large Markdowns
If the parsed JSON is too large for the LLM context window:
1. Send section headings + first 2 bullets per section as a summary
2. Let the LLM create the storyline from the summary
3. Then do a second pass: for each slide, send the full content of its source sections and ask the LLM to distill it into the slide's content fields

### Retry Logic
- If the LLM returns invalid JSON, retry up to 3 times
- If slide count is outside 10-15, re-prompt with explicit correction
- Always validate the output schema before passing to renderer

---

## Stage 3: PPTX Renderer (`renderer/engine.py`)

### Purpose
Take the slide blueprint JSON + template PPTX and produce the final presentation.

### Library
```bash
pip install python-pptx
```

### Template Layout Mapping

The three provided templates have these layouts (indices may vary slightly):

| Index | Layout Name   | Use For                |
|-------|---------------|------------------------|
| 0     | Cover / 1_Cover | Title slide           |
| 1     | Divider       | Section break slides   |
| 2     | Blank         | Content slides (programmatic) |
| 3     | Title only    | Content with header placeholder |
| 4     | Thank You     | Closing slide          |

**IMPORTANT:** At startup, the renderer must introspect the template to find layout indices by name, NOT hardcode indices. Different templates may have layouts in different order.

```python
def get_layout_by_name(prs, name_contains):
    """Find a layout whose name contains the given string (case-insensitive)."""
    for layout in prs.slide_layouts:
        if name_contains.lower() in layout.name.lower():
            return layout
    return prs.slide_layouts[2]  # fallback to Blank
```

### Rendering Loop

```python
from pptx import Presentation

def render(blueprint, template_path, output_path):
    prs = Presentation(template_path)
    
    for slide_data in blueprint["slides"]:
        slide_type = slide_data["type"]
        
        if slide_type == "cover":
            render_cover(prs, slide_data)
        elif slide_type == "section_divider":
            render_divider(prs, slide_data)
        elif slide_type == "thank_you":
            render_thank_you(prs, slide_data)
        elif slide_type == "chart":
            render_chart_slide(prs, slide_data)
        elif slide_type == "table":
            render_table_slide(prs, slide_data)
        else:
            render_content_slide(prs, slide_data)
    
    # Remove the original template slides (they're empty placeholders)
    remove_template_slides(prs)
    
    prs.save(output_path)
```

### Design Constants (`config.py`)

```python
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor

# Slide dimensions (standard widescreen 13.33 x 7.5 inches)
SLIDE_WIDTH = Inches(13.33)
SLIDE_HEIGHT = Inches(7.5)

# Margins
MARGIN_LEFT = Inches(0.6)
MARGIN_RIGHT = Inches(0.6)
MARGIN_TOP = Inches(0.5)
MARGIN_BOTTOM = Inches(0.5)

# Content area
CONTENT_WIDTH = SLIDE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
CONTENT_TOP = Inches(1.6)  # Below title area

# Typography
TITLE_FONT_SIZE = Pt(32)
SUBTITLE_FONT_SIZE = Pt(18)
BODY_FONT_SIZE = Pt(14)
CAPTION_FONT_SIZE = Pt(11)
CARD_HEADING_SIZE = Pt(18)
STAT_NUMBER_SIZE = Pt(48)
STAT_LABEL_SIZE = Pt(12)

# The Caspr template theme colors (extract from template at runtime)
# Fallback defaults based on analysis of provided templates:
COLOR_PRIMARY = RGBColor(0xE8, 0x3F, 0x33)    # Red accent
COLOR_TEXT_DARK = RGBColor(0x1A, 0x1A, 0x1A)   # Near black
COLOR_TEXT_LIGHT = RGBColor(0xFF, 0xFF, 0xFF)   # White
COLOR_TEXT_MUTED = RGBColor(0x66, 0x66, 0x66)   # Gray
COLOR_CARD_BG = RGBColor(0xF9, 0xF0, 0xEE)     # Light pink/cream
COLOR_CARD_BORDER = RGBColor(0xE8, 0x3F, 0x33)  # Red border

# Spacing
CARD_GAP = Inches(0.3)
ELEMENT_GAP = Inches(0.3)
INNER_PADDING = Inches(0.2)

# Chart colors (theme-consistent)
CHART_COLORS = [
    RGBColor(0xE8, 0x3F, 0x33),  # Primary red
    RGBColor(0x1A, 0x1A, 0x1A),  # Black
    RGBColor(0xF9, 0xCB, 0xC2),  # Light pink
    RGBColor(0x66, 0x66, 0x66),  # Gray
    RGBColor(0xE8, 0x7D, 0x73),  # Medium red
]
```

### Layout Rendering Functions (`renderer/layouts.py`)

Each layout function receives a `slide` object and `slide_data` dict, and places shapes at precise coordinates.

#### Key Principles
1. **Grid system** — all elements snap to a consistent grid
2. **No text walls** — max 6 bullets per area, each under 15 words
3. **Clear hierarchy** — title (32pt bold) > heading (18pt bold) > body (14pt regular) > caption (11pt muted)
4. **Consistent spacing** — 0.3" between elements, 0.6" margins
5. **Visual variety** — don't use the same layout for consecutive slides
6. **Slide numbers** — add to bottom-right of every content slide

#### Layout: two_column
```
┌─────────────────────────────────────────┐
│ Title (32pt bold)                        │
│─────────────────────────────────────────│
│ ┌──────────────┐  ┌──────────────┐      │
│ │  Left Column │  │ Right Column │      │
│ │  Heading     │  │  Heading     │      │
│ │  • Point 1   │  │  • Point 1   │      │
│ │  • Point 2   │  │  • Point 2   │      │
│ │  • Point 3   │  │  • Point 3   │      │
│ └──────────────┘  └──────────────┘      │
└─────────────────────────────────────────┘
```

#### Layout: three_cards
```
┌─────────────────────────────────────────┐
│ Title (32pt bold)                        │
│─────────────────────────────────────────│
│ ┌──────────┐ ┌──────────┐ ┌──────────┐ │
│ │ 01       │ │ 02       │ │ 03       │ │
│ │ Heading  │ │ Heading  │ │ Heading  │ │
│ │ • Point  │ │ • Point  │ │ • Point  │ │
│ │ • Point  │ │ • Point  │ │ • Point  │ │
│ └──────────┘ └──────────┘ └──────────┘ │
└─────────────────────────────────────────┘
```

#### Layout: key_stats
```
┌─────────────────────────────────────────┐
│ Title (32pt bold)                        │
│─────────────────────────────────────────│
│                                          │
│   ┌───────┐   ┌───────┐   ┌───────┐    │
│   │ 6.8x  │   │ $47B  │   │  70%  │    │
│   │ growth │   │ spent │   │ share │    │
│   └───────┘   └───────┘   └───────┘    │
│                                          │
└─────────────────────────────────────────┘
```

#### Layout: timeline
```
┌─────────────────────────────────────────┐
│ Title (32pt bold)                        │
│─────────────────────────────────────────│
│                                          │
│  (01)────────(02)────────(03)────(04)   │
│   │            │            │       │    │
│  Text 1     Text 2      Text 3  Text 4  │
│                                          │
└─────────────────────────────────────────┘
```

### Chart Generation (`renderer/charts.py`)

Use python-pptx's built-in chart support. This is a huge advantage — native charts look professional and are editable in PowerPoint.

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

def add_bar_chart(slide, data, position):
    chart_data = CategoryChartData()
    chart_data.categories = data["categories"]
    for series in data["series"]:
        chart_data.add_series(series["name"], series["values"])
    
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        position.left, position.top,
        position.width, position.height,
        chart_data
    ).chart
    
    # Style the chart
    chart.has_legend = len(data["series"]) > 1
    plot = chart.plots[0]
    plot.gap_width = 80
    # Apply theme colors to series
```

Supported chart types:
- `bar` → `XL_CHART_TYPE.COLUMN_CLUSTERED`
- `pie` → `XL_CHART_TYPE.PIE`
- `line` → `XL_CHART_TYPE.LINE_MARKERS`
- `area` → `XL_CHART_TYPE.AREA`

### Table Generation (`renderer/tables.py`)

```python
def add_styled_table(slide, table_data, position):
    rows = len(table_data["rows"]) + 1  # +1 for header
    cols = len(table_data["headers"])
    
    table_shape = slide.shapes.add_table(
        rows, cols,
        position.left, position.top,
        position.width, position.height
    )
    table = table_shape.table
    
    # Style header row
    for i, header in enumerate(table_data["headers"]):
        cell = table.cell(0, i)
        cell.text = header
        # Bold, theme color background
    
    # Fill data rows with alternating background
    for r, row in enumerate(table_data["rows"]):
        for c, value in enumerate(row):
            cell = table.cell(r + 1, c)
            cell.text = str(value)
```

### Infographic Generation (`renderer/infographics.py`)

Build infographics using python-pptx shapes:

1. **Process flow** — Rounded rectangles connected by arrow connectors
2. **Timeline** — Circles with numbers + vertical lines + text boxes
3. **Comparison** — Two colored columns with pros/cons
4. **Icon list** — Colored circles (numbered) + heading + description rows

All infographics must use only programmatic shapes. NO external images.

### Removing Template Placeholder Slides

The template files come with 5-6 placeholder slides. After adding all content slides, remove the originals:

```python
def remove_template_slides(prs):
    """Remove the original template slides (first N slides that were in the template)."""
    # Track how many slides were in the template originally
    # Delete them by XML manipulation
    for i in range(template_slide_count):
        rId = prs.slides._sldIdLst[0].get('r:id')
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]
```

---

## Slide Flow (Mandatory Order)

Every generated presentation MUST follow this structure:

1. **Title Slide** (Cover layout) — always slide 1
2. **Agenda** (optional) — table of contents if markdown has one
3. **Executive Summary** — if markdown has one
4. **Section content slides** — the bulk of the presentation (various layouts)
5. **Chart/data slides** — interspersed where relevant
6. **Conclusion / Key Takeaways** — second-to-last
7. **Thank You** (Thank You layout) — always last slide

---

## Critical Quality Rules

### Visual (30% of judging score)
- NO walls of text — max 6 bullet points per slide, 15 words per bullet
- Clear font hierarchy — title 32pt > heading 18pt > body 14pt > caption 11pt
- Consistent 0.6" margins from all slide edges
- 0.3" minimum gap between all elements
- Cards and shapes must have consistent corner radius
- Use template's color palette — don't introduce random colors
- Every slide needs at least one visual element (shape, chart, table, card)
- Vary layouts across slides — never use same layout twice in a row
- Left-align body text, center only titles

### Charts & Tables (15% of judging score)
- Auto-detect numerical data and generate appropriate chart type
- Charts must be native python-pptx charts (editable in PowerPoint)
- Tables must have styled headers and alternating row colors
- Chart colors must match the template theme
- Add chart title and axis labels

### Content Coverage (15% of judging score)
- Every major section from the markdown must appear
- Executive summary must be included if present in markdown
- No important data or insight should be dropped
- Maintain the narrative flow of the original document

### Code Quality (30% of judging score)
- Clean, modular, well-documented code
- Each module has a single responsibility
- Type hints on all functions
- Docstrings on all public functions
- Error handling with graceful fallbacks
- Logging (use Python `logging` module)

---

## Dependencies

```
# requirements.txt
mistune>=3.0.0
python-pptx>=0.6.23
groq>=0.4.0
python-dotenv>=1.0.0
```

---

## Error Handling & Fallbacks

- **LLM returns invalid JSON** → Retry 3 times with progressively explicit prompts. If all fail, fall back to a rule-based storyline generator that creates one slide per H2 section.
- **Template has unexpected layouts** → Log warning, fall back to Blank layout for all content slides.
- **Markdown has no tables/numbers** → Skip chart slides gracefully, fill with more content slides.
- **Markdown is very short** → Generate minimum 10 slides by expanding content (add agenda, add section dividers).
- **Markdown is very long** → The LLM must aggressively summarize. Cap at 15 slides.
- **Missing executive summary** → Skip that slide, adjust numbering.
- **Any exception during slide rendering** → Catch, log, skip that slide, continue with the rest. Never crash.

---

## Testing

Run against ALL provided test case markdowns before submission:

```bash
# Run against all test cases
for md in test_cases/*.md; do
    python main.py --md "$md" --template templates/Template_AI_Bubble.pptx --output "outputs/$(basename $md .md).pptx"
done
```

Verify each output:
1. Opens without errors in LibreOffice/PowerPoint/Google Slides
2. Has 10-15 slides
3. Has proper title slide and thank you slide
4. No overlapping text or elements
5. Charts render correctly (if applicable)
6. Text is readable (not cut off, not overflowing)

---

## Common Mistakes to Avoid

These are from the hackathon's "Common Mistakes" guide:

1. **Text overflow** — text going outside shape boundaries. Always set `word_wrap = True` on text frames.
2. **Overlapping shapes** — calculate positions carefully. Test with long text.
3. **Inconsistent spacing** — use constants from config.py, not magic numbers.
4. **Missing slide numbers** — add to every content slide (bottom-right).
5. **Wall of text slides** — the LLM prompt must enforce content limits.
6. **Charts without labels** — always add title, axis labels, legend.
7. **Ignoring the template** — all slides MUST use layouts from the provided template.
8. **Random colors** — only use colors from the template theme.
9. **Same layout repeated** — vary layouts to keep the deck visually interesting.
10. **Elements outside margins** — nothing within 0.5" of slide edges.

---

## Implementation Priority

Build in this order:
1. `main.py` — CLI skeleton with argparse
2. `parser/md_parser.py` — get structured JSON from any markdown
3. `storyline/prompts.py` + `storyline/generator.py` — LLM integration
4. `renderer/engine.py` — basic render loop (cover + blank slides with titles)
5. `renderer/layouts.py` — two_column, three_cards, key_stats, single_focus
6. `renderer/charts.py` — bar, pie, line charts
7. `renderer/tables.py` — styled tables
8. `renderer/infographics.py` — timeline, process_flow
9. `config.py` — extract theme colors from template at runtime
10. Error handling + edge cases + testing against all test cases
