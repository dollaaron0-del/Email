"""Lokales Sprachmodell über Ollama - Klassifizierung, Spracherkennung, Phishing-Prüfung."""
from __future__ import annotations

import json
import re

import requests

from ..log import log
from ..models import Mail
from . import CATEGORIES, parse_llm_json


def _chat(cfg: dict, prompt: str, *, json_format: bool, temperature: float, timeout: int, http=None) -> str:
    url = cfg.get("ollama_url") or "http://localhost:11434/api/chat"
    body = {
        "model": cfg.get("ollama_model"),
        "messages": [{"role": "user", "content": prompt}],
        "stream": False,
        "options": {"temperature": temperature},
    }
    if json_format:
        body["format"] = "json"
    resp = (http or requests).post(url, json=body, timeout=timeout)
    resp.raise_for_status()
    return json.loads(resp.content.decode("utf-8"))["message"]["content"]


def tags_url(cfg: dict) -> str:
    return (cfg.get("ollama_url") or "http://localhost:11434/api/chat").replace("/api/chat", "/api/tags")


def classify(mail: Mail, cfg: dict, hotels_block: str, examples_block: str, http=None) -> dict:
    """Kategorie + Hotel bestimmen. Wirft nie - bei Fehlern Kategorie 'unknown'."""

    def run(body_text: str) -> dict:
        prompt = f"""Klassifiziere diese E-Mail im Kontext Hotelmanagement.

Kategorien (waehle EXAKT eine):
- booking: Buchungsanfragen, Zimmerverfuegbarkeit, Reservierungen, Gruppenanfragen - AUCH von privaten Adressen wie gmail.com, hotmail.com etc.
- application: Bewerbungen, Lebenslauf-Einsendungen, Jobanfragen - erkennbar an Karrieredomains (stepstone, linkedin, hotelpersonal, dki-lehrlinge) ODER expliziten Bewerbungsformulierungen
- newsletter: Regelmaessige Rundschreiben, Magazine, Branchennews, oft mit Abmelde-Link
- advertisement: Werbung, Verkaufsangebote, B2B-Marketing, Dienstleistungsangebote, Produktpraesentation
- sonstiges: Wichtige Geschaeftskorrespondenz, Partneranfragen, Kooperationen, externe Emails die in keine andere Kategorie passen

Hotels:
{hotels_block}
{examples_block}
Email:
Von: {mail.sender_email} | Betreff: {mail.subject}
Text: {body_text}

Antworte NUR mit JSON: {{"category":"...","hotel_name":null,"confidence":"high","reason":"..."}}
"""
        r = parse_llm_json(_chat(cfg, prompt, json_format=True, temperature=0.1, timeout=180, http=http))
        cat = r.get("category")
        if cat not in CATEGORIES + ["unknown"]:
            cat = "unknown"
        hotel = r.get("hotel_name")
        result = {
            "category": cat,
            "hotel": "" if hotel in (None, "null") else str(hotel),
            "confidence": str(r.get("confidence") or "").lower(),
            "reason": str(r.get("reason") or ""),
        }
        log.info("  -> %s | Hotel: %s | %s | %s", cat, result["hotel"], result["confidence"], result["reason"])
        return result

    body = mail.body or ""
    try:
        return run(body[:800])  # erst kurz (schneller, weniger Timeouts)
    except Exception:
        log.warning("  -> Erster Versuch fehlgeschlagen, retry mit laengerem Text...")
        try:
            return run(body[:2000])
        except Exception as exc:
            log.error("  -> Klassifizierungsfehler: %s", exc)
            return {"category": "unknown", "hotel": "", "confidence": "low", "reason": "Klassifizierungsfehler"}


def detect_language(mail: Mail, cfg: dict, http=None) -> str:
    snippet = (mail.body or "")[:400]
    prompt = ("Detect the language of this email. Reply with ONLY one of these values: de, en, es\n\n"
              f"Subject: {mail.subject}\nText: {snippet}")
    try:
        text = _chat(cfg, prompt, json_format=False, temperature=0.0, timeout=120, http=http)
        lang = re.sub(r"[^a-z]", "", text.strip().lower())[:2]
        if lang in ("de", "en", "es"):
            return lang
    except Exception:
        pass
    return "de"


def phishing_check(mail: Mail, cfg: dict, http=None) -> dict | None:
    """Phishing-Einschätzung per Ollama. None, wenn Ollama nicht antwortet."""
    snippet = (mail.body or "")[:500]
    prompt = ("Analysiere diese E-Mail auf Phishing, Betrug oder Schadsoftware. Antworte NUR mit JSON.\n\n"
              f"Betreff: {mail.subject}\nAbsender: {mail.sender_email}\nInhalt: {snippet}\n\n"
              'Format: {"is_phishing": false, "confidence": "low", "reason": ""}')
    try:
        r = parse_llm_json(_chat(cfg, prompt, json_format=True, temperature=0.0, timeout=60, http=http))
        return {
            "is_phishing": r.get("is_phishing") is True or str(r.get("is_phishing")).lower() == "true",
            "confidence": str(r.get("confidence") or "low").lower(),
            "reason": str(r.get("reason") or ""),
        }
    except Exception as exc:
        log.warning("  -> Sicherheitsanalyse (LLM) fehlgeschlagen: %s", exc)
        return None
