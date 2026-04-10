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

    Uses a shared pool of Groq clients (round-robin) to distribute API calls
    across multiple API keys, working around free-tier TPM rate limits.
    Set GROQ_API_KEY_1, GROQ_API_KEY_2, … in .env alongside GROQ_API_KEY.
    """

    # Class-level client pool — shared across all BaseAgent instances
    _clients: list[Groq] = []
    _client_index: int = 0
    _pool_initialized: bool = False

    def __init__(self) -> None:
        if not BaseAgent._pool_initialized:
            BaseAgent._clients = self._init_client_pool()
            BaseAgent._pool_initialized = True

    # ------------------------------------------------------------------
    # Client pool (round-robin across multiple API keys)
    # ------------------------------------------------------------------

    @staticmethod
    def _init_client_pool() -> list[Groq]:
        """Initialise Groq clients from all GROQ_API_KEY* env variables.

        Reads GROQ_API_KEY plus GROQ_API_KEY_1 through GROQ_API_KEY_10.
        Deduplicates keys so the same key isn't counted twice.
        Returns empty list when no keys are found.
        """
        keys: list[str] = []

        main_key = os.getenv("GROQ_API_KEY")
        if main_key:
            keys.append(main_key)

        for i in range(1, 11):
            key = os.getenv(f"GROQ_API_KEY_{i}")
            if key and key not in keys:
                keys.append(key)

        if not keys:
            logger.warning("No GROQ_API_KEY* found in environment — agents unavailable.")
            return []

        clients = [Groq(api_key=k) for k in keys]
        logger.info(
            "Groq client pool: %d key(s) loaded for round-robin distribution.",
            len(clients),
        )
        return clients

    def _next_client(self) -> tuple[Groq, int]:
        """Return the next client in round-robin rotation.

        Returns:
            Tuple of (Groq client, 1-based key index for logging).

        Raises:
            ValueError: If the client pool is empty.
        """
        if not BaseAgent._clients:
            raise ValueError(f"{self.__class__.__name__}: No Groq clients available.")
        idx = BaseAgent._client_index % len(BaseAgent._clients)
        BaseAgent._client_index += 1
        return BaseAgent._clients[idx], idx + 1

    @property
    def available(self) -> bool:
        """True when at least one Groq client is ready."""
        return len(BaseAgent._clients) > 0

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

        Automatically rotates through the client pool so consecutive calls
        hit different API keys, spreading the TPM (tokens-per-minute) load.

        Args:
            prompt:     User-turn prompt string.
            system:     System prompt string.
            max_tokens: Maximum tokens in the completion (controls cost/latency).

        Returns:
            Raw response string from the model.

        Raises:
            ValueError: On API error or empty response.
        """
        client, key_num = self._next_client()

        try:
            response = client.chat.completions.create(
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
            raise ValueError(f"Groq API error (key {key_num}): {exc}") from exc

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
