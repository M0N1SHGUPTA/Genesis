"""
main.py — CLI entry point for MD to PPTX converter.

Usage:
    python main.py --md input.md --output output.pptx
    python main.py --md input.md --output output.pptx --slides 12
    python main.py --md input.md --output output.pptx --template templates/custom.pptx
"""

import argparse
import logging
import os
import sys
import time

# ---------------------------------------------------------------------------
# Logging setup — configured before any other imports so that log messages
# from imported modules are captured from the very start of execution.
# Format: HH:MM:SS  LEVEL     message
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    """Define and parse all CLI arguments.

    Returns an argparse.Namespace with attributes:
        md            — path to input markdown file
        output        — path for the output .pptx file
        slides        — optional target slide count (10-15)
        template      — optional explicit template .pptx path
        templates_dir — folder to scan for templates when auto-selecting
        debug         — flag to enable DEBUG-level logging
    """
    parser = argparse.ArgumentParser(
        prog="md-to-pptx",
        description="Convert a Markdown file into a professional PowerPoint presentation.",
    )

    # Required: the source markdown document
    parser.add_argument(
        "--md",
        required=True,
        metavar="FILE",
        help="Path to the input Markdown file (.md)",
    )

    # Required: where to write the finished .pptx
    parser.add_argument(
        "--output",
        required=True,
        metavar="FILE",
        help="Path for the output PowerPoint file (.pptx)",
    )

    # Optional: hint to the LLM about how many slides to produce.
    # When omitted, the LLM decides based on content density (10-15 range).
    parser.add_argument(
        "--slides",
        type=int,
        default=None,
        metavar="N",
        help="Target number of slides (10-15). Auto-chosen when omitted.",
    )

    # Optional: skip auto-selection and use a specific template file.
    # When omitted, the renderer scans --templates-dir and picks the best match.
    parser.add_argument(
        "--template",
        default=None,
        metavar="FILE",
        help=(
            "Path to a specific .pptx Slide Master template. "
            "When omitted the system auto-selects the best template "
            "from the templates/ folder."
        ),
    )

    # Directory that holds all available Slide Master .pptx files.
    # Only used when --template is not explicitly provided.
    parser.add_argument(
        "--templates-dir",
        default="templates",
        metavar="DIR",
        help="Folder that contains Slide Master .pptx files (default: templates/)",
    )

    # Flag: print DEBUG-level messages (stage internals, LLM prompts, etc.)
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable verbose debug logging.",
    )

    return parser.parse_args()


def validate_args(args: argparse.Namespace) -> None:
    """Validate CLI arguments and exit with a helpful message on failure.

    Checks performed:
    - Input markdown file exists on disk
    - Output directory exists (the file itself need not exist yet)
    - --slides is within the 10-15 range when provided
    - Explicit --template file exists when provided
    - --templates-dir exists when auto-selection is needed (no --template given)
    """

    # --- Input markdown ---
    if not os.path.isfile(args.md):
        logger.error("Markdown file not found: %s", args.md)
        sys.exit(1)
    if not args.md.lower().endswith(".md"):
        # Non-fatal: warn and continue — the file might still be valid markdown
        logger.warning("Input file does not have a .md extension: %s", args.md)

    # --- Output path ---
    # We only need the parent directory to exist; the file will be created/overwritten.
    output_dir = os.path.dirname(os.path.abspath(args.output))
    if not os.path.isdir(output_dir):
        logger.error("Output directory does not exist: %s", output_dir)
        sys.exit(1)
    if not args.output.lower().endswith(".pptx"):
        logger.warning("Output file does not have a .pptx extension: %s", args.output)

    # --- Slide count range ---
    # The LLM is constrained to 10-15 slides; reject out-of-range hints early.
    if args.slides is not None and not (10 <= args.slides <= 15):
        logger.error("--slides must be between 10 and 15, got %d", args.slides)
        sys.exit(1)

    # --- Explicit template (optional) ---
    if args.template is not None:
        if not os.path.isfile(args.template):
            logger.error("Template file not found: %s", args.template)
            sys.exit(1)
        if not args.template.lower().endswith(".pptx"):
            logger.warning("Template file does not have a .pptx extension: %s", args.template)

    # --- Templates directory (needed only when auto-selecting) ---
    # If the user did not provide --template, the renderer must be able to scan
    # the templates directory to find a suitable Slide Master.
    if args.template is None and not os.path.isdir(args.templates_dir):
        logger.error(
            "Templates directory not found: %s  "
            "(create it and place .pptx Slide Master files inside, "
            "or pass --template to specify one explicitly)",
            args.templates_dir,
        )
        sys.exit(1)


def run_pipeline(args: argparse.Namespace) -> None:
    """Execute the three-stage conversion pipeline.

    Stages:
        1. Parse   — markdown file → structured JSON (sections, bullets, tables, numbers)
        2. Storyline — structured JSON → slide blueprint JSON (via Groq LLM)
        3. Render  — slide blueprint + template → output .pptx file

    Imports are deferred to inside this function so that argument validation
    (which runs before this) never fails due to a missing dependency.
    """

    start = time.time()  # track total wall-clock time for the run

    # ------------------------------------------------------------------
    # Stage 1 — Parse markdown into structured JSON
    # Extracts: title, subtitle, executive summary, sections, subsections,
    # bullets, paragraphs, tables, and numerical data blocks.
    # ------------------------------------------------------------------
    logger.info("Stage 1/3 — Parsing markdown: %s", args.md)
    from parser.md_parser import MarkdownParser  # noqa: PLC0415

    parser_obj = MarkdownParser()
    parsed = parser_obj.parse(args.md)  # returns a dict matching the schema in CLAUDE.md
    logger.info(
        "Parsed %d section(s), %d table(s), %d numerical block(s)",
        parsed.get("total_sections", 0),
        parsed.get("total_tables", 0),
        parsed.get("total_numerical_blocks", 0),
    )

    # ------------------------------------------------------------------
    # Stage 2 — Generate slide blueprint via LLM
    # The LLM decides: slide count, layout per slide, which data becomes
    # a chart, content trimming, and narrative flow.
    # ------------------------------------------------------------------
    logger.info("Stage 2/3 — Generating slide storyline via LLM …")
    from storyline.generator import StorylineGenerator  # noqa: PLC0415

    # target_slides=None lets the LLM decide freely within the 10-15 range
    generator = StorylineGenerator(target_slides=args.slides)
    blueprint = generator.generate(parsed)  # returns slide blueprint dict
    logger.info(
        "Blueprint ready: %d slide(s) planned",
        blueprint.get("total_slides", "?"),
    )

    # ------------------------------------------------------------------
    # Stage 3 — Render blueprint + template → .pptx
    # The renderer auto-selects a template from templates_dir when
    # template_path is None, otherwise uses the explicitly provided file.
    # ------------------------------------------------------------------
    logger.info("Stage 3/3 — Rendering presentation …")
    from renderer.engine import Renderer  # noqa: PLC0415

    renderer = Renderer(
        template_path=args.template,    # None → auto-select from templates_dir
        templates_dir=args.templates_dir,
    )
    renderer.render(blueprint, parsed, args.output)

    elapsed = time.time() - start
    logger.info("Done! Output saved to: %s  (%.1fs)", args.output, elapsed)


def main() -> None:
    """Program entry point — parse args, validate, then run the pipeline."""

    args = parse_args()

    # Elevate log level to DEBUG before anything else runs, so all
    # subsequent module-level log calls are captured at full verbosity.
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
        logger.debug("Debug logging enabled")

    # Fail fast: check all inputs before doing any expensive work
    validate_args(args)

    try:
        run_pipeline(args)
    except KeyboardInterrupt:
        # Ctrl-C: exit cleanly without a traceback
        logger.warning("Interrupted by user.")
        sys.exit(130)  # 130 is the conventional exit code for SIGINT
    except Exception as exc:  # noqa: BLE001
        # Any unhandled exception from the pipeline: log with full traceback
        # and exit with a non-zero code so the caller knows something failed.
        logger.exception("Fatal error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
