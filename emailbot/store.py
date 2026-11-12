"""Warteschlange, Statistik und Lernbeispiele - gleiche Dateien und Formate wie die PowerShell-Version.

Bot und Dashboard laufen im selben Prozess; alle Lese-Ändern-Schreiben-Vorgänge auf einer Datei
laufen unter derselben Sperre (pro Dateipfad), damit sich die beiden nicht gegenseitig überschreiben.
"""
from __future__ import annotations

import json
import threading
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator

from .config import read_json, write_json, write_text_atomic

_LOCKS: dict[str, threading.RLock] = {}
_LOCKS_GUARD = threading.Lock()


def file_lock(path: Path) -> threading.RLock:
    key = str(Path(path).resolve()).lower()
    with _LOCKS_GUARD:
        if key not in _LOCKS:
            _LOCKS[key] = threading.RLock()
        return _LOCKS[key]


def _flatten(obj: Any, out: list) -> None:
    """Alte PowerShell-Dateien enthalten teils verschachtelte {"value": [...]}-Strukturen."""
    if isinstance(obj, list):
        for item in obj:
            _flatten(item, out)
    elif isinstance(obj, dict):
        if "id" in obj:
            out.append(obj)
        elif "value" in obj:
            _flatten(obj["value"], out)


class QueueStore:
    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self.lock = file_lock(self.path)

    def load(self) -> list[dict]:
        with self.lock:
            flat: list[dict] = []
            _flatten(read_json(self.path, default=[]), flat)
            seen: set[str] = set()
            result = []
            for entry in flat:
                if entry.get("id") and entry["id"] not in seen:
                    seen.add(entry["id"])
                    result.append(entry)
            return result

    def save(self, entries: list[dict]) -> None:
        with self.lock:
            parts = [json.dumps(e, ensure_ascii=False, separators=(",", ":")) for e in entries if e is not None]
            write_text_atomic(self.path, "[]\n" if not parts else "[\n" + ",\n".join(parts) + "\n]\n")

    @contextmanager
    def edit(self) -> Iterator[list[dict]]:
        with self.lock:
            entries = self.load()
            yield entries
            self.save(entries)

    def add(self, entry: dict) -> None:
        with self.edit() as entries:
            entries.append(entry)

    def get(self, entry_id: str) -> dict | None:
        return next((e for e in self.load() if e.get("id") == entry_id), None)

    def update(self, entry_id: str, fields: dict) -> bool:
        with self.edit() as entries:
            for entry in entries:
                if entry.get("id") == entry_id:
                    entry.update(fields)
                    return True
        return False


STAT_CATEGORIES = ["newsletter", "application", "advertisement", "booking", "sonstiges", "unknown", "error"]


class StatsStore:
    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self.lock = file_lock(self.path)

    def load(self) -> dict:
        with self.lock:
            raw = read_json(self.path, default=None) or {}
            cats = raw.get("categories") or {}
            return {
                "total_processed": int(raw.get("total_processed") or 0),
                "today_processed": int(raw.get("today_processed") or 0),
                "last_date": str(raw.get("last_date") or ""),
                "last_checked": str(raw.get("last_checked") or "-"),
                "categories": {k: int(cats.get(k) or 0) for k in STAT_CATEGORIES} | {
                    k: int(v or 0) for k, v in cats.items() if k not in STAT_CATEGORIES},
                "hotels": {k: int(v or 0) for k, v in (raw.get("hotels") or {}).items()},
                "recent": list(raw.get("recent") or []),
            }

    def save(self, stats: dict) -> None:
        with self.lock:
            write_json(self.path, stats)

    def set_last_checked(self) -> None:
        with self.lock:
            s = self.load()
            s["last_checked"] = datetime.now().strftime("%d.%m.%Y %H:%M")
            self.save(s)

    def record(self, entry: dict, result: str = "done") -> None:
        if result != "done":
            return
        with self.lock:
            s = self.load()
            today = datetime.now().strftime("%Y-%m-%d")
            if s["last_date"] != today:
                s["today_processed"] = 0
                s["last_date"] = today
            s["total_processed"] += 1
            s["today_processed"] += 1
            cat = entry.get("category") or "unknown"
            s["categories"][cat] = s["categories"].get(cat, 0) + 1
            if cat == "booking" and entry.get("hotel"):
                s["hotels"][entry["hotel"]] = s["hotels"].get(entry["hotel"], 0) + 1
            rec = {
                "time": datetime.now().strftime("%d.%m. %H:%M"),
                "subject": entry.get("subject"),
                "sender": entry.get("sender_name"),
                "category": cat,
                "hotel": entry.get("hotel"),
                "folder": entry.get("target_folder"),
                "corrected": bool(entry.get("is_corrected")),
                "auto": bool(entry.get("auto_approved")),
            }
            s["recent"] = ([rec] + s["recent"])[:100]
            self.save(s)


class CorrectionsStore:
    """Lernbeispiele aus deinen Korrekturen (werden der KI als Beispiele mitgegeben)."""

    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self.lock = file_lock(self.path)

    def load(self) -> list[dict]:
        with self.lock:
            data = read_json(self.path, default=[])
            flat: list[dict] = []
            for item in data if isinstance(data, list) else [data]:
                if isinstance(item, dict):
                    flat.append(item)
            return flat

    def save_correction(self, entry: dict) -> None:
        with self.lock:
            sender = entry.get("sender") or ""
            domain = sender.split("@", 1)[1] if "@" in sender else sender
            new = {
                "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "subject": entry.get("subject"),
                "sender_domain": domain,
                "wrong_category": entry.get("category"),
                "correct_category": entry.get("correct_category"),
                "correct_hotel": entry.get("correct_hotel") or "",
            }
            all_items = ([new] + self.load())[:50]
            parts = [json.dumps(c, ensure_ascii=False, separators=(",", ":")) for c in all_items]
            write_text_atomic(self.path, "[\n" + ",\n".join(parts) + "\n]\n")

    def examples_block(self) -> str:
        """Max. 2 Beispiele pro Muster (falsch->richtig), höchstens 15 insgesamt."""
        seen: dict[str, int] = {}
        examples = []
        for c in self.load():
            key = f"{c.get('wrong_category')}->{c.get('correct_category')}"
            if seen.get(key, 0) < 2:
                examples.append(c)
                seen[key] = seen.get(key, 0) + 1
            if len(examples) >= 15:
                break
        if not examples:
            return ""
        lines = []
        for c in examples:
            hotel = f" -> Hotel: {c['correct_hotel']}" if c.get("correct_hotel") else ""
            lines.append(f'  - Betreff: "{c.get("subject")}" | Domain: {c.get("sender_domain")} | '
                         f'War: {c.get("wrong_category")} | Richtig: {c.get("correct_category")}{hotel}')
        return "\nLernbeispiele (HOHE PRIORITAET - diese Muster wurden korrigiert):\n" + "\n".join(lines) + "\n"
