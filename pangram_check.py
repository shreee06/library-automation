"""
pangram_check.py

Runs extracted paper text through Pangram's AI-detection and plagiarism
APIs and returns a clean, flat dict ready for the Excel logger.

Setup:
    pip install pangram-sdk
    export PANGRAM_API_KEY=your_key_here   (or put it in a .env file)

NOTE: Pangram's exact response field names may differ slightly from what's
used below -- the public docs we had access to while building this were a
quickstart summary, not the full field reference. The first time you run
this against a real API key, print(result) once (see the __main__ block)
and adjust the field names in `_safe_get` calls below to match what you
actually get back. That's a 5-minute fix, not a redesign.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any

from pathlib import Path

from dotenv import load_dotenv
from pangram import Pangram

# Load .env from this file's folder, regardless of the working directory.
load_dotenv(Path(__file__).resolve().parent / ".env")


@dataclass
class CheckResult:
    ai_percent: float | None
    human_percent: float | None
    ai_tier: str | None
    plagiarism_detected: bool | None
    plagiarism_match_count: int | None
    raw_predict: dict[str, Any]
    raw_plagiarism: dict[str, Any]


def _get_client() -> Pangram:
    api_key = os.environ.get("PANGRAM_API_KEY")
    if not api_key:
        raise RuntimeError(
            "PANGRAM_API_KEY is not set. Put it in a .env file "
            "(see .env.example) or export it in your shell."
        )
    return Pangram(api_key=api_key)


def _safe_get(d: dict, *keys: str, default=None):
    """Try several possible key names in order, return the first hit."""
    for k in keys:
        if k in d:
            return d[k]
    return default


def check_paper(text: str) -> CheckResult:
    """
    Run both the AI-detection and plagiarism checks on a block of text.
    """
    client = _get_client()

    predict_result = client.predict(text)
    plagiarism_result = client.check_plagiarism(text)

    # Pangram returns fractions (0-1); the log wants percentages.
    ai_fraction = _safe_get(predict_result, "fraction_ai")
    human_fraction = _safe_get(predict_result, "fraction_human")
    ai_percent = round(ai_fraction * 100, 1) if ai_fraction is not None else None
    if ai_percent is None:
        # Fail loudly rather than logging a blank row as "Processed".
        raise ValueError(
            "Could not find the AI score in Pangram's predict() response "
            f"(keys seen: {sorted(predict_result)}). "
            "Update the field names in pangram_check.py."
        )
    human_percent = (
        round(human_fraction * 100, 1)
        if human_fraction is not None
        else round(100 - ai_percent, 1)
    )
    ai_tier = _safe_get(predict_result, "headline", "prediction_short")

    matches = _safe_get(plagiarism_result, "plagiarized_content", default=[])
    plagiarism_match_count = len(matches) if matches else 0
    plagiarism_detected = _safe_get(plagiarism_result, "plagiarism_detected")
    if plagiarism_detected is None:
        # Keep the two plagiarism columns consistent if the flag is absent.
        plagiarism_detected = plagiarism_match_count > 0

    return CheckResult(
        ai_percent=ai_percent,
        human_percent=human_percent,
        ai_tier=ai_tier,
        plagiarism_detected=plagiarism_detected,
        plagiarism_match_count=plagiarism_match_count,
        raw_predict=predict_result,
        raw_plagiarism=plagiarism_result,
    )


if __name__ == "__main__":
    # Quick manual test: python pangram_check.py "some text to check"
    import sys

    if len(sys.argv) != 2:
        print("Usage: python pangram_check.py \"text to check\"")
        sys.exit(1)

    result = check_paper(sys.argv[1])
    print("AI %:", result.ai_percent)
    print("Human %:", result.human_percent)
    print("Tier:", result.ai_tier)
    print("Plagiarism detected:", result.plagiarism_detected)
    print("Plagiarism matches:", result.plagiarism_match_count)
    print()
    print("Raw predict() response (check this to fix field names above):")
    print(result.raw_predict)
    print()
    print("Raw check_plagiarism() response:")
    print(result.raw_plagiarism)
