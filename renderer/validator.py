"""
renderer/validator.py — Python Design Enforcer.

A purely rule-based validator that runs AFTER the agent pipeline and BEFORE
the renderer. It catches structural problems that the LLM might introduce
and silently fixes them so the renderer never sees broken data.

Why Python instead of an LLM:
  - These are deterministic rules (text length, array emptiness, slide ordering)
  - An LLM cannot "see" whether shapes will overflow — Python can count words
  - No API calls → zero latency, zero cost, never fails

Rules enforced:
  1. First slide must be "cover"
  2. Last slide must be "thank_you"
  3. Every slide must have a non-empty "title"
  4. Cover subtitle ≤ 20 words
  5. All bullet / point text ≤ 15 words
  6. points / cards / items / stats / steps arrays must not be empty
  7. three_cards must have exactly 3 cards
  8. key_stats must have 2–4 stats
  9. Slide numbers renumbered sequentially (1, 2, 3, …)
  10. total_slides updated to match the actual slide count
  11. Layout variety: no two consecutive content slides may use the same layout
"""

from __future__ import annotations

import logging

logger = logging.getLogger(__name__)

# These layouts can be swapped when consecutive duplicates are found
_LAYOUT_ROTATION = [
    "single_focus", "two_column", "three_cards",
    "icon_list", "key_stats", "timeline",
]

# Slide types that have a "layout" field subject to variety rules
_CONTENT_TYPES = {"content", "executive_summary", "conclusion"}


class DesignEnforcer:
    """Apply design rules to a blueprint dict before it reaches the renderer.

    All methods mutate the blueprint in-place and log warnings for each
    fix applied. The enforcer never raises — if a rule cannot be applied
    cleanly, it logs and moves on.

    Usage:
        enforcer = DesignEnforcer()
        blueprint = enforcer.enforce(blueprint)
    """

    def enforce(self, blueprint: dict) -> dict:
        """Apply all design rules to the blueprint.

        Args:
            blueprint: Slide blueprint dict from the agent pipeline.

        Returns:
            The same dict, mutated to satisfy all design rules.
        """
        try:
            slides = blueprint.get("slides", [])
            if not isinstance(slides, list):
                blueprint["slides"] = []
                return blueprint

            self._ensure_cover(blueprint)
            self._ensure_thank_you(blueprint)

            for slide in blueprint["slides"]:
                self._fix_title(slide)
                self._fix_cover_subtitle(slide)
                self._fix_bullet_lengths(slide)
                self._fix_empty_arrays(slide)
                self._fix_three_cards_count(slide)
                self._fix_key_stats_count(slide)

            self._fix_layout_variety(blueprint["slides"])
            self._renumber(blueprint["slides"])
            blueprint["total_slides"] = len(blueprint["slides"])

        except Exception as exc:
            logger.error("DesignEnforcer encountered unexpected error: %s", exc)

        return blueprint

    # ------------------------------------------------------------------
    # Rule 1 & 2: Cover first, Thank You last
    # ------------------------------------------------------------------

    def _ensure_cover(self, blueprint: dict) -> None:
        """Inject a cover slide at position 0 if one is missing."""
        slides = blueprint["slides"]
        if not slides or slides[0].get("type") != "cover":
            logger.warning("Validator: first slide is not 'cover' — injecting.")
            slides.insert(0, {
                "slide_number": 1,
                "type": "cover",
                "title": blueprint.get("presentation_title", "Presentation"),
                "subtitle": "",
            })

    def _ensure_thank_you(self, blueprint: dict) -> None:
        """Append a thank_you slide if one is missing at the end."""
        slides = blueprint["slides"]
        if not slides or slides[-1].get("type") != "thank_you":
            logger.warning("Validator: last slide is not 'thank_you' — appending.")
            slides.append({
                "slide_number": len(slides) + 1,
                "type": "thank_you",
                "title": "Thank You",
                "subtitle": "",
            })

    # ------------------------------------------------------------------
    # Rule 3: Every slide needs a title
    # ------------------------------------------------------------------

    def _fix_title(self, slide: dict) -> None:
        """Replace a missing or empty title with a safe default."""
        if not slide.get("title"):
            slide["title"] = slide.get("type", "slide").replace("_", " ").title()
            logger.debug("Validator: auto-filled empty title on slide %s.", slide.get("slide_number"))

    # ------------------------------------------------------------------
    # Rule 4: Cover subtitle ≤ 20 words
    # ------------------------------------------------------------------

    def _fix_cover_subtitle(self, slide: dict) -> None:
        """Truncate cover subtitle to 20 words if needed."""
        if slide.get("type") != "cover":
            return
        subtitle = slide.get("subtitle", "")
        if not subtitle:
            return
        words = subtitle.split()
        if len(words) > 20:
            slide["subtitle"] = " ".join(words[:20]) + "…"
            logger.debug("Validator: truncated cover subtitle to 20 words.")

    # ------------------------------------------------------------------
    # Rule 5: Bullet text ≤ 15 words
    # ------------------------------------------------------------------

    def _fix_bullet_lengths(self, slide: dict) -> None:
        """Trim any bullet / point string that exceeds 15 words."""
        _truncate = lambda t: " ".join(t.split()[:15]) + ("…" if len(t.split()) > 15 else "")

        # top-level "points" (agenda, single_focus, conclusion)
        if "points" in slide:
            slide["points"] = [_truncate(p) for p in slide["points"]]

        # two_column / comparison / executive_summary: left.points / right.points
        for side in ("left", "right"):
            col = slide.get(side)
            if isinstance(col, dict) and "points" in col:
                col["points"] = [_truncate(p) for p in col["points"]]

        # three_cards: cards[*].points
        for card in slide.get("cards", []):
            if "points" in card:
                card["points"] = [_truncate(p) for p in card["points"]]

        # timeline / process_flow: steps[*].description
        for step in slide.get("steps", []):
            if "description" in step:
                step["description"] = _truncate(step["description"])

        # icon_list: items[*].description
        for item in slide.get("items", []):
            if "description" in item:
                item["description"] = _truncate(item["description"])

        # single_focus: "focus" field
        if "focus" in slide:
            words = slide["focus"].split()
            if len(words) > 20:
                slide["focus"] = " ".join(words[:20]) + "…"

    # ------------------------------------------------------------------
    # Rule 6: No empty arrays
    # ------------------------------------------------------------------

    def _fix_empty_arrays(self, slide: dict) -> None:
        """Replace empty content arrays with a safe fallback value."""
        _PLACEHOLDER = "See document for details."

        # points
        if "points" in slide and not slide["points"]:
            slide["points"] = [_PLACEHOLDER]
            logger.debug("Validator: filled empty 'points' on slide %s.", slide.get("slide_number"))

        # two_column / comparison sides
        for side in ("left", "right"):
            col = slide.get(side)
            if isinstance(col, dict):
                if not col.get("points"):
                    col["points"] = [_PLACEHOLDER]
                if not col.get("heading"):
                    col["heading"] = side.title()

        # cards
        for card in slide.get("cards", []):
            if not card.get("points"):
                card["points"] = [_PLACEHOLDER]
            if not card.get("heading"):
                card["heading"] = f"Card {card.get('number', '')}"

        # steps
        if "steps" in slide and not slide["steps"]:
            slide["steps"] = [{"number": "01", "heading": "Step 1", "description": _PLACEHOLDER}]
            logger.debug("Validator: filled empty 'steps' on slide %s.", slide.get("slide_number"))

        # items
        if "items" in slide and not slide["items"]:
            slide["items"] = [{"number": "01", "heading": "Item 1", "description": _PLACEHOLDER}]
            logger.debug("Validator: filled empty 'items' on slide %s.", slide.get("slide_number"))

        # stats
        if "stats" in slide and not slide["stats"]:
            slide["stats"] = [{"value": "—", "label": "Key metric"}]
            logger.debug("Validator: filled empty 'stats' on slide %s.", slide.get("slide_number"))

    # ------------------------------------------------------------------
    # Rule 7: three_cards needs exactly 3 cards
    # ------------------------------------------------------------------

    def _fix_three_cards_count(self, slide: dict) -> None:
        """Pad or trim the cards array to exactly 3 entries."""
        if slide.get("layout") != "three_cards":
            return
        cards = slide.get("cards", [])
        # Trim excess
        slide["cards"] = cards[:3]
        # Pad to 3
        while len(slide["cards"]) < 3:
            n = len(slide["cards"]) + 1
            slide["cards"].append({
                "number": str(n).zfill(2),
                "heading": f"Point {n}",
                "points": ["See document for details."],
            })
            logger.debug(
                "Validator: padded three_cards to 3 cards on slide %s.",
                slide.get("slide_number"),
            )

    # ------------------------------------------------------------------
    # Rule 8: key_stats needs 2–4 stats
    # ------------------------------------------------------------------

    def _fix_key_stats_count(self, slide: dict) -> None:
        """Ensure key_stats has at least 2 and at most 4 stat entries."""
        if slide.get("layout") != "key_stats":
            return
        stats = slide.get("stats", [])
        slide["stats"] = stats[:4]
        while len(slide["stats"]) < 2:
            slide["stats"].append({"value": "—", "label": "Key metric"})
            logger.debug(
                "Validator: padded key_stats on slide %s.",
                slide.get("slide_number"),
            )

    # ------------------------------------------------------------------
    # Rule 11: No consecutive content slides with the same layout
    # ------------------------------------------------------------------

    def _fix_layout_variety(self, slides: list[dict]) -> None:
        """Ensure no two consecutive content slides share the same layout.

        When a collision is detected, the second slide's layout is replaced
        with the next layout in _LAYOUT_ROTATION that doesn't match either
        neighbour.

        Args:
            slides: The full slides list (mutated in-place).
        """
        for i in range(1, len(slides)):
            curr = slides[i]
            prev = slides[i - 1]

            if curr.get("type") not in _CONTENT_TYPES:
                continue
            if prev.get("type") not in _CONTENT_TYPES:
                continue

            curr_layout = curr.get("layout", "")
            prev_layout = prev.get("layout", "")

            if curr_layout and curr_layout == prev_layout:
                # Find a replacement that differs from both neighbours
                next_layout = prev.get("layout", "") if i < len(slides) - 1 else ""
                for candidate in _LAYOUT_ROTATION:
                    if candidate != curr_layout and candidate != next_layout:
                        logger.debug(
                            "Validator: slide %s layout changed from '%s' to '%s' (consecutive duplicate).",
                            curr.get("slide_number"),
                            curr_layout,
                            candidate,
                        )
                        curr["layout"] = candidate
                        # Content fields for the new layout may be missing — fill placeholders
                        self._add_layout_placeholders(curr, candidate)
                        break

    def _add_layout_placeholders(self, slide: dict, layout: str) -> None:
        """Add the minimum required content fields for a layout if they're absent.

        Called when the layout is swapped to avoid rendering the slide with
        missing fields.

        Args:
            slide:  Slide dict (mutated in-place).
            layout: The new layout name.
        """
        _PH = "See document for details."

        if layout == "single_focus":
            slide.setdefault("focus", slide.get("title", ""))
            slide.setdefault("points", [_PH])

        elif layout in ("two_column", "comparison"):
            slide.setdefault("left", {"heading": "Overview", "points": [_PH]})
            slide.setdefault("right", {"heading": "Details", "points": [_PH]})

        elif layout == "three_cards":
            if not slide.get("cards"):
                slide["cards"] = [
                    {"number": str(i + 1).zfill(2), "heading": f"Point {i+1}", "points": [_PH]}
                    for i in range(3)
                ]

        elif layout == "key_stats":
            slide.setdefault("stats", [{"value": "—", "label": "Key metric"}])

        elif layout in ("timeline", "process_flow"):
            slide.setdefault("steps", [{"number": "01", "heading": "Step 1", "description": _PH}])

        elif layout == "icon_list":
            slide.setdefault("items", [{"number": "01", "heading": "Item 1", "description": _PH}])

    # ------------------------------------------------------------------
    # Rule 9: Sequential slide numbers
    # ------------------------------------------------------------------

    @staticmethod
    def _renumber(slides: list[dict]) -> None:
        """Assign consecutive slide_number values starting at 1."""
        for i, slide in enumerate(slides):
            slide["slide_number"] = i + 1
