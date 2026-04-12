# Genesis — Markdown to PowerPoint Pipeline

> **Hackathon Submission | EZ Labs | April 2026**

Genesis converts research-heavy Markdown documents into polished, professional `.pptx` slide decks using a multi-agent AI pipeline backed by Groq LLMs — with full graceful fallback when no API key is available.

---

## Table of Contents

1. [What Genesis Does](#what-genesis-does)
2. [System Architecture](#system-architecture)
3. [Key Design Decisions](#key-design-decisions)
4. [Repository Structure](#repository-structure)
5. [Setup Instructions](#setup-instructions)
6. [Run Steps](#run-steps)
7. [Inputs & Outputs](#inputs--outputs)
8. [Fallback Behavior (No API Key)](#fallback-behavior-no-api-key)
9. [Judging Criteria Alignment](#judging-criteria-alignment)

---

## What Genesis Does

Genesis is **not** a 1:1 Markdown renderer. It is an intelligent presentation engine that:

- Parses dense research documents into structured content
- Uses a **3-agent AI pipeline** to plan narrative arc, slide structure, and content transformation
- Renders the result into a `.pptx` using real PowerPoint templates, native charts, native tables, and programmatic layouts
- Produces 10–15 professionally designed slides regardless of how long the input document is

The output looks like a deck a human designer would produce — fewer slides than document sections, a logical story arc, charts only where numeric data exists, and visual layouts that adapt to content density.

---

## System Architecture

### Pipeline Overview

```
Markdown file (.md)
        │
        ▼
   main.py  ──── CLI orchestrator, validates args, runs all stages
        │
        ▼
parser/md_parser.py  ──── MarkdownParser.parse()
        │
        ▼
   Parsed document dict
        │
        ▼
agents/pipeline.py  ──── AgentPipeline.generate()
        │
        ├──▶  Agent 1: ContentExtractor   (agents/content_extractor.py)
        │         Distils sections into presentation-friendly insights,
        │         identifies chart/table/KPI opportunities
        │
        ├──▶  Agent 2: StorylinePlanner   (agents/storyline_planner.py)
        │         Plans slide sequence, assigns types and layouts,
        │         enforces 10–15 slide budget
        │
        └──▶  Agent 3: ContentTransformer (agents/content_transformer.py)
                  Writes final slide-ready content for every slide
                  in the approved plan
        │
        ▼
renderer/validator.py  ──── DesignEnforcer
        │         Validates blueprint: word limits, layout variety,
        │         required slides, safe defaults
        ▼
   Blueprint dict (renderer-ready)
        │
        ▼
renderer/engine.py  ──── Renderer.render()
        │
        ├── Template selection + theme colour extraction
        ├── Layout renderers  (renderer/layouts.py)
        ├── Native chart rendering  (renderer/charts.py)
        ├── Styled table rendering  (renderer/tables.py)
        └── Shape / infographic primitives  (renderer/infographics.py)
        │
        ▼
   outputs/<filename>.pptx
```

### Internal Data Representations

Genesis passes through four clean internal representations:

| Stage | Representation | Description |
|---|---|---|
| 1 | `parsed` | Raw structural facts extracted from the Markdown |
| 2 | `extracted_content` | Presentation-friendly insights, visual opportunities, KPI stats |
| 3 | `slide_plan` | Slide sequence, types, layout assignments |
| 4 | `blueprint` | Final renderer-ready payload per slide |

---

## Key Design Decisions

### 1. Separation of Concerns Across Four Stages

Each stage solves a fundamentally different problem:

- **Parsing** (`parser/`) is about **correctness** — what is in the document?
- **Agent planning** (`agents/`) is about **judgment** — what matters and how should the deck flow?
- **Validation** (`renderer/validator.py`) is about **safety rails** — are presentation constraints met before rendering starts?
- **Rendering** (`renderer/`) is about **geometry** — how do we draw this into PowerPoint?

This split means each module can be extended, tested, or replaced independently.

### 2. Three Narrow Agents Instead of One Monolithic Prompt

Instead of asking a single LLM prompt to "turn this document into a presentation", Genesis breaks the task into three agents with narrower, clearly scoped responsibilities:

- **ContentExtractor**: What are the insights? What should be visualised?
- **StorylinePlanner**: What is the deck structure? (does **not** write content)
- **ContentTransformer**: Write the actual slide copy for the approved plan.

Narrowing each agent's task reduces hallucination, improves JSON reliability, and makes failure more recoverable.

### 3. Graceful Fallback Without LLM

Every agent has a deterministic rule-based fallback. If Groq is unavailable or rate-limited, the entire pipeline still runs end-to-end using Python heuristics. The deck may be less nuanced but the system **never crashes due to LLM unavailability**.

### 4. Round-Robin API Key Pooling

`agents/base_agent.py` supports up to 11 Groq API keys (`GROQ_API_KEY`, `GROQ_API_KEY_1` … `GROQ_API_KEY_10`) and distributes calls round-robin. This allows sustained throughput on large documents without hitting per-key rate limits.

### 5. Template-Driven Visual Quality

The renderer extracts colour themes from the chosen `.pptx` template at runtime and overwrites global config values. Charts, tables, cards, and accents all follow the template's palette automatically — so visual quality scales with the template without code changes.

### 6. Per-Slide Fault Isolation

The renderer wraps each slide's rendering in `try/except`. A single bad slide logs a warning and is skipped; the rest of the deck is still produced. This prevents one edge-case from killing the entire output.

### 7. Blueprint Validation Layer

`renderer/validator.py` (DesignEnforcer) sits between the agents and the renderer to enforce:

- word count limits per slide
- layout variety (avoids repetitive slide types)
- presence of mandatory slides (cover, conclusion, thank-you)
- safe defaults for missing fields

This is done in Python rather than via prompt text because code constraints are deterministic.

---

## Repository Structure

```
Genesis/
├── main.py                     # CLI entry point and stage orchestrator
├── config.py                   # Global slide dimensions, colours, fonts, spacing
├── restyle.py                  # One-off post-processing script (not part of main flow)
├── CLAUDE.md                   # Older project notes (useful context, partially outdated)
│
├── parser/
│   ├── __init__.py
│   └── md_parser.py            # Markdown → parsed dict
│
├── agents/
│   ├── __init__.py
│   ├── base_agent.py           # Groq client pooling, retries, JSON parsing
│   ├── content_extractor.py    # Agent 1 — insight extraction
│   ├── storyline_planner.py    # Agent 2 — slide structure planning
│   ├── content_transformer.py  # Agent 3 — slide copy generation
│   └── pipeline.py             # Active Stage 2 orchestrator
│
├── storyline/                  # Legacy single-agent architecture (not active)
│   ├── __init__.py
│   ├── generator.py
│   └── prompts.py
│
├── renderer/
│   ├── __init__.py
│   ├── engine.py               # Stage 3 entry point, template theming
│   ├── validator.py            # Blueprint repair and enforcement
│   ├── layouts.py              # Content layout renderers
│   ├── visuals.py              # Reusable visual primitives
│   ├── charts.py               # Native Office chart rendering
│   ├── tables.py               # Styled table rendering
│   ├── infographics.py         # Complex shape-based visual helpers
│   └── utils.py                # Shared low-level rendering utilities
│
├── templates/                  # PowerPoint master templates
│   ├── master_template_1.pptx
│   ├── master_template_2.pptx
│   └── master_template_3.pptx
│
├── test_cases/                 # Example Markdown inputs
│   ├── accenture.md
│   ├── AI_Bubble.md
│   └── UAE.md
│
├── outputs/                    # Generated .pptx files land here
└── target/                     # Reference design assets (not in runtime flow)
```

---

## Setup Instructions

### Prerequisites

- Python 3.10+ (project was developed on Python 3.14)
- A Groq API key (free tier works; optional — fallback runs without it)

### 1. Clone the Repository

```bash
git clone <your-repo-url>
cd Genesis
```

### 2. Create and Activate a Virtual Environment

**Windows (PowerShell):**
```powershell
python -m venv env
.\env\Scripts\Activate.ps1
```

**macOS / Linux:**
```bash
python -m venv env
source env/bin/activate
```

### 3. Install Dependencies

```bash
pip install python-pptx groq python-dotenv mistune lxml
```

### 4. Configure Environment Variables

Create a `.env` file in the project root:

```env
GROQ_API_KEY=your_primary_key_here

# Optional — for round-robin rate-limit spreading on large documents
GROQ_API_KEY_1=optional_key_2
GROQ_API_KEY_2=optional_key_3
```

> Only `GROQ_API_KEY` is required. The numbered variants are optional.

---

## Run Steps

### Basic Run

```powershell
python main.py --md test_cases\UAE.md --output outputs\uae_demo.pptx
```

### Specify Slide Count (10–15)

```powershell
python main.py --md test_cases\accenture.md --output outputs\accenture_demo.pptx --slides 12
```

### Use an Explicit Template

```powershell
python main.py --md test_cases\AI_Bubble.md --template templates\master_template_2.pptx --output outputs\ai_bubble_demo.pptx
```

### Enable Debug Logging

```powershell
python main.py --md test_cases\UAE.md --output outputs\uae_debug.pptx --debug
```

### All CLI Arguments

| Argument | Required | Description |
|---|---|---|
| `--md` | ✅ Yes | Path to input Markdown file |
| `--output` | ✅ Yes | Output `.pptx` path |
| `--slides` | No | Target slide count (10–15) |
| `--template` | No | Explicit `.pptx` template path |
| `--templates-dir` | No | Directory for auto template selection (default: `templates/`) |
| `--debug` | No | Enables verbose logging |

---

## Inputs & Outputs

### Inputs

| Type | Example |
|---|---|
| Markdown file | `test_cases/accenture.md` |
| Markdown file | `test_cases/AI_Bubble.md` |
| Markdown file | `test_cases/UAE.md` |
| Template | `templates/master_template_1.pptx` |
| Template | `templates/master_template_2.pptx` |
| Template | `templates/master_template_3.pptx` |

### Output

All generated decks land in `outputs/`. The renderer appends newly created slides to the chosen template, then removes the original template/demo slides so only the generated content remains.

---

## Fallback Behavior (No API Key)

If Groq is unavailable or no key is set:

| Component | Fallback |
|---|---|
| `ContentExtractor` | Rule-based extraction from bullets, sentences, tables, numeric blocks |
| `StorylinePlanner` | Heuristic slide plan from section headings and content density |
| `ContentTransformer` | Rule-based slide copy from extracted insights |
| `DesignEnforcer` | Runs normally — Python constraints are not LLM-dependent |
| `Renderer` | Runs normally — produces a valid `.pptx` |

**The pipeline is fully operational without any LLM access.** The deck will be less narratively nuanced but structurally correct and visually consistent with the template.

---

## Judging Criteria Alignment

| Criterion | How Genesis Addresses It |
|---|---|
| **Visual Quality (30%)** | Template-driven rendering; theme colours extracted at runtime; native Office charts and tables; multiple layout types; per-slide fault isolation keeps the deck clean |
| **Code Quality & Agentic Development (30%)** | Three-agent pipeline with clear separation of concerns; shared base agent infrastructure; round-robin API key pooling; full fallback chain; validator between agent and renderer |
| **Chart & Table Generation (15%)** | `renderer/charts.py` produces native Office charts; `renderer/tables.py` produces styled tables; `ContentExtractor` identifies visual opportunities per section |
| **Content Coverage (15%)** | Multi-pass chunked extraction for large documents; all H2 sections processed; executive summary, KPIs, and conclusion captured |
| **Innovation (10%)** | Graceful LLM-free fallback; blueprint validation layer between agents and renderer; round-robin key pooling; template-driven palette inheritance |

---

*Built for the EZ Labs Hackathon — April 2026*
