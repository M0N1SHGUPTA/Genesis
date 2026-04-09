"""
md_parser.py — Stage 1: Markdown file → structured dict.

Uses mistune 3.x with renderer='ast' and the table plugin.
The output dict is consumed by storyline/generator.py in Stage 2.
"""

from __future__ import annotations

import logging
import os
import re
from typing import Any

import mistune

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
_CHUNK_SIZE = 256 * 1024        # 256 KB per read chunk
_MAX_FILE_SIZE = 5 * 1024 * 1024  # 5 MB hard cap


# ---------------------------------------------------------------------------
# Public class
# ---------------------------------------------------------------------------

class MarkdownParser:
    """Parse a Markdown file into a structured dict for the storyline generator.

    Usage:
        parser = MarkdownParser()
        result = parser.parse("report.md")
    """

    def __init__(self) -> None:
        # Single reusable mistune instance; renderer='ast' returns list[dict]
        self._md = mistune.create_markdown(renderer="ast", plugins=["table"])

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def parse(self, filepath: str) -> dict:
        """Parse a Markdown file and return its structured representation.

        Args:
            filepath: Path to the .md file.

        Returns:
            Dict with keys: title, subtitle, executive_summary, sections,
            total_sections, total_tables, total_numerical_blocks.

        Raises:
            FileNotFoundError: If filepath does not exist.
        """
        if not os.path.isfile(filepath):
            raise FileNotFoundError(f"Markdown file not found: {filepath}")

        text = self._read_file(filepath)
        tokens = self._get_ast(text)
        filename_stem = os.path.splitext(os.path.basename(filepath))[0]
        doc = self._walk_tokens(tokens, filename_stem)
        return self._finalize(doc)

    # ------------------------------------------------------------------
    # File reading
    # ------------------------------------------------------------------

    def _read_file(self, filepath: str) -> str:
        """Read file in chunks, capping at 5 MB."""
        chunks: list[str] = []
        bytes_read = 0
        with open(filepath, "r", encoding="utf-8", errors="replace") as fh:
            while bytes_read < _MAX_FILE_SIZE:
                chunk = fh.read(_CHUNK_SIZE)
                if not chunk:
                    break
                chunks.append(chunk)
                bytes_read += len(chunk.encode("utf-8"))
        if bytes_read >= _MAX_FILE_SIZE:
            logger.warning("File exceeds 5 MB — content truncated.")
        return "".join(chunks)

    # ------------------------------------------------------------------
    # AST generation
    # ------------------------------------------------------------------

    def _get_ast(self, text: str) -> list[dict]:
        """Run mistune and return the flat block-level token list."""
        try:
            tokens = self._md(text)
            return tokens if isinstance(tokens, list) else []
        except Exception as exc:
            logger.error("mistune parse failed: %s", exc)
            return []

    # ------------------------------------------------------------------
    # Core state-machine walk
    # ------------------------------------------------------------------

    def _walk_tokens(self, tokens: list[dict], filename_stem: str) -> dict:
        """Single linear pass over the token list; builds the output dict."""

        # Initialise the document skeleton
        doc: dict[str, Any] = {
            "title": filename_stem,   # fallback if no H1 found
            "subtitle": "",
            "executive_summary": "",
            "sections": [],
        }

        # State
        current_section: dict | None = None
        current_subsection: dict | None = None
        found_h1 = False
        found_subtitle = False
        in_exec_summary = False
        exec_summary_parts: list[str] = []
        pending_section_content = ""
        # True once ANY H2/H3 has been encountered after the H1. Used to stop
        # the subtitle-auto-capture from stealing paragraphs that belong to
        # the first real section (e.g. the Executive Summary).
        past_top_matter = False
        # True when we're inside a Table-of-Contents or References H2 that
        # we've chosen to skip. Any H3 encountered while this is set must
        # also be suppressed (no implicit section promotion).
        in_skipped_h2 = False

        # ------------------------------------------------------------------
        # Helper: flush subsection into current section
        # ------------------------------------------------------------------
        def flush_subsection() -> None:
            nonlocal current_subsection
            if current_subsection is not None and current_section is not None:
                current_section["subsections"].append(current_subsection)
                current_subsection = None

        # Helper: flush section into doc
        def flush_section() -> None:
            nonlocal current_section, pending_section_content
            if current_section is not None:
                current_section["content"] = pending_section_content.strip()
                doc["sections"].append(current_section)
                current_section = None
                pending_section_content = ""

        # ------------------------------------------------------------------
        # Token dispatch
        # ------------------------------------------------------------------
        for token in tokens:
            t = token.get("type", "")

            # Skip irrelevant token types
            if t in ("blank_line", "thematic_break", "block_code", "block_html"):
                continue

            # ---- Headings ------------------------------------------------
            if t == "heading":
                level = token.get("attrs", {}).get("level", 1)
                heading_text = self._extract_text(token).strip()

                if level == 1 and not found_h1:
                    doc["title"] = heading_text
                    found_h1 = True
                    in_exec_summary = False

                elif level == 2:
                    flush_subsection()
                    flush_section()
                    past_top_matter = True
                    # Skip Table of Contents and References sections entirely.
                    _low = heading_text.lower().strip().lstrip("0123456789. ")
                    if (
                        _low.startswith("table of contents")
                        or _low.startswith("contents")
                        or _low.startswith("references")
                        or _low.startswith("bibliography")
                        or _low.startswith("appendix")
                    ):
                        in_exec_summary = False
                        in_skipped_h2 = True
                        current_section = None
                        current_subsection = None
                        pending_section_content = ""
                        continue
                    in_skipped_h2 = False
                    in_exec_summary = "executive summary" in heading_text.lower()
                    if not in_exec_summary:
                        current_section = self._new_section(heading_text)
                    pending_section_content = ""

                elif level == 3:
                    # Any H3 under a skipped H2 must also be suppressed.
                    if in_skipped_h2:
                        continue
                    # H3 before any H2 and no subtitle yet → treat as subtitle.
                    if current_section is None and not found_subtitle and not past_top_matter:
                        doc["subtitle"] = heading_text
                        found_subtitle = True
                        continue
                    flush_subsection()
                    in_exec_summary = False
                    past_top_matter = True
                    if current_section is None:
                        # H3 before any H2 (but subtitle already taken) —
                        # create an implicit section.
                        current_section = self._new_section(heading_text)
                    current_subsection = self._new_subsection(heading_text, 3)

                else:
                    # H4/H5 — flatten into current subsection as a labelled bullet
                    if current_subsection is not None:
                        current_subsection["bullets"].append(f"[{heading_text}]")
                    elif current_section is not None:
                        current_subsection = self._new_subsection(heading_text, level)

            # ---- Paragraphs ----------------------------------------------
            elif t == "paragraph":
                text = self._extract_text(token).strip()
                if not text:
                    continue
                if in_skipped_h2:
                    continue

                if (
                    found_h1
                    and not found_subtitle
                    and current_section is None
                    and not in_exec_summary
                    and not past_top_matter
                ):
                    doc["subtitle"] = text
                    found_subtitle = True

                elif in_exec_summary:
                    exec_summary_parts.append(text)

                elif current_subsection is not None:
                    current_subsection["paragraphs"].append(text)
                    nums = self._extract_numerical(text)
                    current_subsection["numerical_data"].extend(nums)

                elif current_section is not None:
                    # Paragraph between H2 and first H3 → section intro text
                    pending_section_content += (" " + text)
                    nums = self._extract_numerical(text)
                    # Attach to a lazy implicit subsection if numerical
                    if nums:
                        if current_subsection is None:
                            current_subsection = self._new_subsection("", 3)
                        current_subsection["numerical_data"].extend(nums)

            # ---- Lists ---------------------------------------------------
            elif t == "list":
                if in_skipped_h2:
                    continue
                bullets = self._extract_bullets(token)
                if not bullets:
                    continue

                if in_exec_summary:
                    exec_summary_parts.extend(bullets)
                elif current_subsection is not None:
                    current_subsection["bullets"].extend(bullets)
                elif current_section is not None:
                    # List before first H3 — create implicit subsection
                    if current_subsection is None:
                        current_subsection = self._new_subsection("", 3)
                    current_subsection["bullets"].extend(bullets)

            # ---- Tables --------------------------------------------------
            elif t == "table":
                table_data = self._extract_table(token)
                if table_data is None:
                    continue

                if current_subsection is None and current_section is not None:
                    current_subsection = self._new_subsection("", 3)

                if current_subsection is not None:
                    current_subsection["tables"].append(table_data)

            # ---- Block quotes — treat as paragraphs ----------------------
            elif t == "block_quote":
                text = self._extract_text(token).strip()
                if text and current_subsection is not None:
                    current_subsection["paragraphs"].append(text)

            else:
                logger.debug("Skipping unsupported token type: %s", t)

        # End of token loop — flush whatever is still open
        flush_subsection()
        flush_section()
        doc["executive_summary"] = " ".join(exec_summary_parts).strip()

        return doc

    # ------------------------------------------------------------------
    # Text extraction helpers
    # ------------------------------------------------------------------

    def _extract_text(self, node: dict) -> str:
        """Recursively extract plain text from any AST node.

        - Links: keep link text, drop URL
        - Images: keep alt text, drop URL
        - Code spans: dropped
        - Soft/hard breaks: replaced with a space
        """
        t = node.get("type", "")

        if t == "text":
            return node.get("raw", "")
        if t in ("softbreak", "linebreak"):
            return " "
        if t == "codespan":
            return ""
        if t in ("link", "image"):
            # Recurse into children for visible text / alt text
            return "".join(self._extract_text(c) for c in node.get("children", []))

        children = node.get("children", [])
        if children:
            return "".join(self._extract_text(c) for c in children)

        # Leaf node with raw content (fallback)
        return node.get("raw", "")

    def _extract_bullets(self, node: dict) -> list[str]:
        """Extract bullet strings from a list token (handles nesting)."""
        bullets: list[str] = []
        for item in node.get("children", []):
            if item.get("type") != "list_item":
                continue
            for child in item.get("children", []):
                if child.get("type") == "list":
                    # Nested list — flatten
                    bullets.extend(self._extract_bullets(child))
                else:
                    text = self._extract_text(child).strip()
                    if text:
                        bullets.append(text)
        return bullets

    def _extract_table(self, node: dict) -> dict | None:
        """Parse a table token into {headers, rows}.

        Returns None if the table is malformed or empty.
        """
        try:
            children = node.get("children", [])
            head = next((c for c in children if c.get("type") == "table_head"), None)
            body = next((c for c in children if c.get("type") == "table_body"), None)

            if not head:
                return None

            headers = [
                self._extract_text(cell).strip()
                for cell in head.get("children", [])
            ]
            col_count = len(headers)
            if col_count == 0:
                return None

            rows: list[list[str]] = []
            if body:
                for row_node in body.get("children", []):
                    if row_node.get("type") != "table_row":
                        continue
                    cells = [
                        self._extract_text(c).strip()
                        for c in row_node.get("children", [])
                    ]
                    # Pad short rows, trim long rows
                    while len(cells) < col_count:
                        cells.append("")
                    rows.append(cells[:col_count])

            return {"headers": headers, "rows": rows}

        except Exception as exc:
            logger.warning("Skipping malformed table: %s", exc)
            return None

    # ------------------------------------------------------------------
    # Numerical data extraction
    # ------------------------------------------------------------------

    def _extract_numerical(self, text: str) -> list[dict]:
        """Extract numerical data points from a plain-text string.

        Looks for:
        1. Year-keyed series:  "2020: $10B, 2021: $25B"
        2. Label-value pairs:  "Market share: 70%"

        Returns a list of {"context": str, "values": dict[str, float]}.
        """
        entries: list[dict] = []

        # --- Pattern 1: year-keyed series ---
        year_pattern = re.compile(r'\b(20\d{2}|19\d{2})\b')
        years = year_pattern.findall(text)
        if len(years) >= 2:
            values: dict[str, float] = {}
            # Split on year tokens to find adjacent numbers
            segments = year_pattern.split(text)
            current_year: str | None = None
            for seg in segments:
                if year_pattern.fullmatch(seg):
                    current_year = seg
                elif current_year:
                    m = re.search(r'[\$€£]?\s*(-?\d[\d,]*(?:\.\d+)?)', seg)
                    if m:
                        try:
                            values[current_year] = float(m.group(1).replace(",", ""))
                            current_year = None
                        except ValueError:
                            pass
            if len(values) >= 2:
                ctx_m = re.match(r'^([A-Za-z][^\d:,\n]{3,40})', text)
                context = ctx_m.group(1).strip().rstrip(":") if ctx_m else "data"
                entries.append({"context": context, "values": values})

        # --- Pattern 2: label: value pairs ---
        label_val = re.compile(
            r'([A-Za-z][A-Za-z0-9 \-/]{2,40})'   # label
            r'[:\s]+'
            r'(-?\d[\d,]*(?:\.\d+)?)'              # number
            r'\s*([%$BMKbmk](?:illion)?)?'          # optional unit
        )
        for m in label_val.finditer(text):
            ctx = m.group(1).strip().rstrip(":")
            # Skip pure years already handled above
            if re.fullmatch(r'20\d{2}|19\d{2}', ctx):
                continue
            try:
                val = float(m.group(2).replace(",", ""))
            except ValueError:
                continue
            entries.append({"context": ctx, "values": {"value": val}})

        return entries

    # ------------------------------------------------------------------
    # Factory helpers
    # ------------------------------------------------------------------

    def _new_section(self, heading: str) -> dict:
        return {"heading": heading, "level": 2, "content": "", "subsections": []}

    def _new_subsection(self, heading: str, level: int) -> dict:
        return {
            "heading": heading,
            "level": level,
            "bullets": [],
            "paragraphs": [],
            "tables": [],
            "has_numerical_data": False,
            "numerical_data": [],
        }

    # ------------------------------------------------------------------
    # Finalise
    # ------------------------------------------------------------------

    def _finalize(self, doc: dict) -> dict:
        """Compute counters, set flags, fill in defaults."""

        total_tables = 0
        total_numerical = 0

        for sec in doc["sections"]:
            for sub in sec["subsections"]:
                total_tables += len(sub["tables"])
                total_numerical += len(sub["numerical_data"])
                sub["has_numerical_data"] = bool(sub["numerical_data"])

        doc["total_sections"] = len(doc["sections"])
        doc["total_tables"] = total_tables
        doc["total_numerical_blocks"] = total_numerical

        doc.setdefault("title", "Untitled")
        doc.setdefault("subtitle", "")
        doc.setdefault("executive_summary", "")

        return doc
