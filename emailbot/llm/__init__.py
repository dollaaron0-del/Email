"""Sprachmodelle: Langdock (Cloud) und Ollama (lokal)."""
from __future__ import annotations

import json

CATEGORIES = ["booking", "application", "newsletter", "advertisement", "sonstiges"]
CONFIDENCES = ["high", "medium", "low"]


def parse_llm_json(text: str) -> dict:
    """Toleranter JSON-Parser: entfernt ```json-Blöcke und Text um das JSON herum."""
    t = (text or "").strip()
    start, end = t.find("{"), t.rfind("}")
    if start < 0 or end <= start:
        raise ValueError(f"Antwort enthaelt kein JSON: {t[:120]}")
    return json.loads(t[start:end + 1])
