"""
agents/base_agent.py — Shared foundation for all three LLM agents.

Provides:
  - Groq client initialisation from GROQ_API_KEY env variable
  - _call_llm()        : single API call with configurable token budget
  - _parse_json()      : robust JSON extractor (handles code fences, prefix text)
  - _run_with_retry()  : up to MAX_RETRIES attempts with progressive correction

All three agents inherit from BaseAgent so none of them duplicate this logic.
"""

from __future__ import annotations

import json
import logging
import os
import re

from dotenv import load_dotenv
from groq import Groq

load_dotenv()

logger = logging.getLogger(__name__)

MODEL = "llama-3.3-70b-versatile"
MAX_RETRIES = 3


class BaseAgent:
    """Base class for all LLM agents.

    Subclasses implement their own public method (e.g. extract(), plan(),
    transform()) and call _run_with_retry() or _call_llm() internally.
    """

    def __init__(self) -> None:
        self._client = self._init_client()

    # ------------------------------------------------------------------
    # Client
    # ------------------------------------------------------------------

    def _init_client(self) -> Groq | None:
        """Initialise Groq client from GROQ_API_KEY.

        Returns None when the key is missing so callers can fall back
        to rule-based logic gracefully.
        """
        api_key = os.getenv("GROQ_API_KEY")
        if not api_key:
            logger.warning(
                "%s: GROQ_API_KEY not set — agent unavailable.",
                self.__class__.__name__,
            )
            return None
        return Groq(api_key=api_key)

    @property
    def available(self) -> bool:
        """True when the Groq client is ready to make API calls."""
        return self._client is not None

    # ------------------------------------------------------------------
    # LLM call
    # ------------------------------------------------------------------

    def _call_llm(
        self,
        prompt: str,
        system: str,
        max_tokens: int = 2048,
    ) -> str:
        """Send a single prompt to the Groq API and return the raw text.

        Args:
            prompt:     User-turn prompt string.
            system:     System prompt string.
            max_tokens: Maximum tokens in the completion (controls cost/latency).

        Returns:
            Raw response string from the model.

        Raises:
            ValueError: On API error or empty response.
        """
        if self._client is None:
            raise ValueError(f"{self.__class__.__name__}: Groq client not initialised.")

        try:
            response = self._client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.2,      # low → deterministic, structured output
                max_tokens=max_tokens,
            )
            text = response.choices[0].message.content or ""
            if not text.strip():
                raise ValueError("LLM returned an empty response.")
            return text
        except Exception as exc:
            raise ValueError(f"Groq API error: {exc}") from exc

    # ------------------------------------------------------------------
    # JSON parsing
    # ------------------------------------------------------------------

    def _parse_json(self, raw: str) -> dict:
        """Extract and parse a JSON object from the LLM's raw response.

        Handles common LLM quirks:
          - Markdown code fences  (```json ... ```)
          - Leading/trailing prose before/after the JSON object
          - Whitespace padding

        Args:
            raw: Raw string from the LLM.

        Returns:
            Parsed Python dict.

        Raises:
            ValueError: If no valid JSON object can be found.
        """
        text = raw.strip()

        # Strip markdown code fences
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
        text = text.strip()

        # Try direct parse
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass

        # Find outermost { ... } block
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            try:
                return json.loads(text[start : end + 1])
            except json.JSONDecodeError as exc:
                raise ValueError(f"JSON parse error: {exc}") from exc

        raise ValueError("No JSON object found in LLM response.")

    # ------------------------------------------------------------------
    # Retry wrapper
    # ------------------------------------------------------------------

    def _run_with_retry(
        self,
        prompt: str,
        system: str,
        max_tokens: int = 2048,
    ) -> dict:
        """Call the LLM up to MAX_RETRIES times, parsing JSON each time.

        On JSON parse failure the same prompt is retried (the temperature is
        already low, so the model often self-corrects on retry). On API errors
        (e.g. rate-limit 429) the same behaviour applies — retry up to the cap.

        Args:
            prompt:     User-turn prompt.
            system:     System prompt.
            max_tokens: Token budget for the completion.

        Returns:
            Parsed Python dict.

        Raises:
            ValueError: If all retries are exhausted.
        """
        last_error = ""
        for attempt in range(1, MAX_RETRIES + 1):
            logger.info(
                "%s — attempt %d/%d",
                self.__class__.__name__,
                attempt,
                MAX_RETRIES,
            )
            try:
                raw = self._call_llm(prompt, system, max_tokens)
                return self._parse_json(raw)
            except (ValueError, KeyError) as exc:
                last_error = str(exc)
                logger.warning(
                    "%s — attempt %d failed: %s",
                    self.__class__.__name__,
                    attempt,
                    last_error,
                )

        raise ValueError(
            f"{self.__class__.__name__}: all {MAX_RETRIES} attempts failed. "
            f"Last error: {last_error}"
        )
