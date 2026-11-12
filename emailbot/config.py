"""Konfiguration und Datei-Hilfen.

Alle Laufzeitdateien liegen im Programmordner (eine Ebene über dem Paket) - dieselben
Dateien wie bei der PowerShell-Version, damit beide Versionen austauschbar sind.
"""
from __future__ import annotations

import copy
import json
import os
import time
from pathlib import Path
from typing import Any

PACKAGE_DIR = Path(__file__).resolve().parent
DEFAULT_BASE_DIR = PACKAGE_DIR.parent

DEFAULTS: dict[str, Any] = {
    "hotels": [],
    "applications_email": "",
    "ollama_model": "llama3.2",
    "ollama_url": "http://localhost:11434/api/chat",
    "check_interval_minutes": 5,
    "scan_max_age_days": 1,
    "scan_max_per_run": 50,
    "skip_domains": [],
    "skip_senders": [],
    "test_mode": False,
    "mail_backend": "outlook",
    "dashboard_port": 5000,
    "llm_provider": "ollama",
    "langdock": {
        "model": "",
        "agent_id": "",
        "region": "eu",
        "base_url": "https://api.langdock.com",
        "api_key_env": "LANGDOCK_API_KEY",
        "fallback_to_ollama": True,
        "body_chars": 3000,
        "timeout_sec": 90,
    },
    "auto_approve": {
        "mode": "off",
        "min_confidence": "high",
        "categories": ["advertisement", "newsletter", "application", "booking"],
        "allow_ollama": False,
    },
    "training_data": {
        "enabled": False,
        "dir": "training_data",
        "body_chars": 4000,
        "compare_ollama": False,
    },
    "reply_templates": {},
}


def deep_merge(base: dict, override: dict) -> dict:
    """Legt override über base; verschachtelte Abschnitte werden zusammengeführt."""
    result = copy.deepcopy(base)
    for key, value in (override or {}).items():
        if value is None and isinstance(result.get(key), dict):
            continue  # "langdock": null o.ä. -> Standardwerte behalten
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def read_json(path: Path, default: Any = None) -> Any:
    """Liest JSON - auch Dateien mit UTF-8-BOM, wie PowerShell sie schreibt."""
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def write_text_atomic(path: Path, text: str) -> None:
    """Schreibt erst in eine Temp-Datei und ersetzt dann - so entsteht nie eine halbe Datei."""
    path = Path(path)
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    for attempt in range(5):
        try:
            os.replace(tmp, path)
            return
        except PermissionError:
            # Virenscanner/Indexer halten die Datei unter Windows manchmal kurz fest
            time.sleep(0.2 * (attempt + 1))
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    try:
        os.remove(tmp)
    except OSError:
        pass


def write_json(path: Path, data: Any, indent: int | None = 2) -> None:
    write_text_atomic(path, json.dumps(data, ensure_ascii=False, indent=indent) + "\n")


def load_config(base_dir: Path) -> dict:
    raw = read_json(Path(base_dir) / "config.json")
    if raw is None:
        raise FileNotFoundError(
            f"config.json fehlt in {base_dir} - bitte config.example.json kopieren und anpassen."
        )
    cfg = deep_merge(DEFAULTS, raw)
    if not cfg["auto_approve"].get("categories"):
        cfg["auto_approve"]["categories"] = list(DEFAULTS["auto_approve"]["categories"])
    return cfg
