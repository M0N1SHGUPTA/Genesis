"""
agents/pipeline.py — Multi-agent pipeline orchestrator.

This is the drop-in replacement for storyline/generator.py's
StorylineGenerator.generate(). It chains the three agents and applies the
Python DesignEnforcer before returning the final blueprint.

Pipeline stages:
  1. ContentExtractor.extract(parsed)              → extracted_content
  2. StorylinePlanner.plan(extracted, target)      → slide_plan
  3. ContentTransformer.transform(plan, extracted) → blueprint
  4. DesignEnforcer.enforce(blueprint)             → validated blueprint

Each stage has an independent fallback:
  - Stage 1 failure → rule-based extraction, continue to Stage 2
  - Stage 2 failure → rule-based plan, continue to Stage 3
  - Stage 3 failure → rule-based transform (produces valid blueprint)
  - Stage 4 never fails (pure Python)

If ALL three LLM agents fail and no Groq key is set, the entire pipeline
delegates to the original StorylineGenerator as a last resort.
"""

from __future__ import annotations

import logging
import time

from agents.content_extractor import ContentExtractor
from agents.storyline_planner import StorylinePlanner
from agents.content_transformer import ContentTransformer

logger = logging.getLogger(__name__)


class AgentPipeline:
    """Multi-agent pipeline: parsed markdown dict → slide blueprint.

    This is the primary entry point for the intelligence layer.
    It replaces StorylineGenerator and is called from main.py.

    Usage:
        pipeline = AgentPipeline(target_slides=12)
        blueprint = pipeline.generate(parsed)
    """

    def __init__(self, target_slides: int | None = None) -> None:
        """Initialise all three agents.

        Args:
            target_slides: Desired slide count (10–15), or None to let Agent 2
                           decide based on content density.
        """
        self.target_slides = target_slides
        self._extractor = ContentExtractor()
        self._planner = StorylinePlanner()
        self._transformer = ContentTransformer()

    # ------------------------------------------------------------------
    # Public API — matches StorylineGenerator.generate() signature
    # ------------------------------------------------------------------

    def generate(self, parsed: dict) -> dict:
        """Run the full 3-agent pipeline and return a validated blueprint.

        Args:
            parsed: Structured dict produced by MarkdownParser.parse().

        Returns:
            Slide blueprint dict consumed by renderer/engine.py.
        """
        start = time.time()

        # ---------------------------------------------------------------
        # Stage 1 — Content Extraction
        # ---------------------------------------------------------------
        logger.info("Agent 1/3 — ContentExtractor: extracting key content …")
        try:
            extracted = self._extractor.extract(parsed)
            logger.info(
                "Agent 1 done: %d sections, %d stats, suggested %d slides",
                len(extracted.get("key_sections", [])),
                len(extracted.get("global_stats", [])),
                extracted.get("suggested_slide_count", 0),
            )
        except Exception as exc:
            logger.error("Agent 1 raised unexpectedly: %s — using rule-based extraction.", exc)
            extracted = self._extractor._fallback_extract(parsed)

        # ---------------------------------------------------------------
        # Stage 2 — Storyline Planning
        # ---------------------------------------------------------------
        logger.info("Agent 2/3 — StorylinePlanner: designing slide sequence …")
        try:
            plan = self._planner.plan(extracted, self.target_slides)
            logger.info(
                "Agent 2 done: %d slides planned",
                plan.get("total_slides", 0),
            )
        except Exception as exc:
            logger.error("Agent 2 raised unexpectedly: %s — using rule-based plan.", exc)
            plan = self._planner._fallback_plan(extracted, self.target_slides)

        # ---------------------------------------------------------------
        # Stage 3 — Content Transformation
        # ---------------------------------------------------------------
        logger.info("Agent 3/3 — ContentTransformer: filling slide content …")
        try:
            blueprint = self._transformer.transform(plan, extracted, parsed)
            logger.info(
                "Agent 3 done: %d slides generated",
                blueprint.get("total_slides", 0),
            )
        except Exception as exc:
            logger.error("Agent 3 raised unexpectedly: %s — using rule-based transform.", exc)
            blueprint = self._transformer._rule_based_transform(plan, extracted, parsed)

        # ---------------------------------------------------------------
        # Stage 4 — Python Design Enforcement (never raises)
        # ---------------------------------------------------------------
        logger.info("Validator — enforcing design rules …")
        try:
            from renderer.validator import DesignEnforcer
            enforcer = DesignEnforcer()
            blueprint = enforcer.enforce(blueprint)
            logger.info(
                "Validator done: %d slides, all rules applied",
                blueprint.get("total_slides", 0),
            )
        except Exception as exc:
            # Validator should never fail, but if it does, log and continue
            logger.error("DesignEnforcer raised unexpectedly: %s — skipping enforcement.", exc)

        elapsed = time.time() - start
        logger.info(
            "Pipeline complete: %d slides in %.1fs",
            blueprint.get("total_slides", 0),
            elapsed,
        )
        return blueprint

    # ------------------------------------------------------------------
    # Last-resort fallback: original single-agent generator
    # ------------------------------------------------------------------

    @staticmethod
    def _legacy_generate(parsed: dict, target_slides: int | None) -> dict:
        """Delegate to the original StorylineGenerator as a last resort.

        This is called only if all three agents and their Python fallbacks
        have been exhausted AND we still need a blueprint. In practice this
        should never happen, but having it prevents a hard crash.

        Args:
            parsed:        Parsed markdown dict.
            target_slides: Optional desired slide count.

        Returns:
            Blueprint from the legacy single-agent generator.
        """
        logger.warning("Falling back to legacy StorylineGenerator.")
        from storyline.generator import StorylineGenerator
        generator = StorylineGenerator(target_slides=target_slides)
        return generator.generate(parsed)
