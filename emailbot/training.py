"""Trainingsdaten sammeln - gleiches Format wie die PowerShell-Version.

Eine Zeile pro Ereignis in training_data/emails_JJJJ-MM.jsonl:
  "prediction" = was die KI vorgeschlagen hat (inkl. Mailtext)
  "decision"   = was danach passiert ist (approved, auto_approved, corrected, rejected, ignored)
Zusammenführen mit:  python -m emailbot.export_training
"""
from __future__ import annotations

import json
import threading
from datetime import datetime
from pathlib import Path

from .llm.langdock import PROMPT_VERSION
from .log import log
from .models import Mail
from .security import unsubscribe_info

_WRITE_LOCK = threading.Lock()


def enabled(cfg: dict) -> bool:
    return (cfg.get("training_data") or {}).get("enabled") is True


def data_dir(base_dir: Path, cfg: dict) -> Path:
    d = Path((cfg.get("training_data") or {}).get("dir") or "training_data")
    return d if d.is_absolute() else Path(base_dir) / d


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


class TrainingLog:
    def __init__(self, base_dir: Path) -> None:
        self.base_dir = Path(base_dir)

    def write_event(self, cfg: dict, record: dict) -> None:
        if not enabled(cfg):
            return
        try:
            d = data_dir(self.base_dir, cfg)
            d.mkdir(parents=True, exist_ok=True)
            path = d / f"emails_{datetime.now():%Y-%m}.jsonl"
            line = json.dumps(record, ensure_ascii=False, separators=(",", ":"))
            with _WRITE_LOCK, open(path, "a", encoding="utf-8", newline="\n") as f:
                f.write(line + "\n")
        except Exception as exc:
            log.warning("  -> Trainingsdaten konnten nicht geschrieben werden: %s", exc)

    @staticmethod
    def input_for(mail: Mail, cfg: dict) -> dict:
        """Alles, was ein Modell später als Eingabe sehen soll - vor dem Verschieben erfassen."""
        max_chars = int((cfg.get("training_data") or {}).get("body_chars") or 4000)
        sender = (mail.sender_email or "").lower()
        u = unsubscribe_info(mail.headers)
        return {
            "message_id": mail.message_id,
            "received": mail.received.strftime("%Y-%m-%dT%H:%M:%S"),
            "sender_domain": sender.split("@", 1)[1] if "@" in sender else "",
            "subject": mail.subject or "",
            "body": (mail.body or "").strip()[:max_chars],
            "attachments": list(mail.attachments),
            "has_list_unsubscribe": bool(u["raw"]),
            "one_click_unsubscribe": bool(u["one_click"]),
        }

    def prediction(self, cfg: dict, entry: dict, train_input: dict | None, prediction: dict | None,
                   extra: dict) -> None:
        if not train_input:
            return
        self.write_event(cfg, {
            "event": "prediction",
            "id": entry.get("id"),
            "time": _now(),
            "prompt_version": PROMPT_VERSION,
            "input": train_input,
            "prediction": prediction,
            "queued_category": entry.get("category"),
            "queued_hotel": entry.get("hotel"),
            "proposed_action": entry.get("proposed_action"),
            "auto": {"mode": entry.get("auto_mode"), "eligible": bool(entry.get("auto_eligible")),
                     "approved": bool(entry.get("auto_approved")), "note": entry.get("auto_note")},
            "extra": extra,
        })

    def decision(self, cfg: dict, entry: dict, decision: str, final_category: str = "", final_hotel: str = "",
                 result: str = "") -> None:
        self.write_event(cfg, {
            "event": "decision",
            "id": entry.get("id"),
            "message_id": entry.get("message_id"),
            "time": _now(),
            "decision": decision,
            "final_category": final_category,
            "final_hotel": final_hotel,
            "was_corrected": bool(entry.get("is_corrected")),
            "result": result,
        })
