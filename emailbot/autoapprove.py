"""Automatische Freigabe: darf ein Eintrag ohne manuelle Prüfung ausgeführt werden?

Modi: off (alles manuell), shadow (Testbetrieb: nur markieren), on (sofort ausführen).
"""
from __future__ import annotations

CONF_RANK = {"low": 1, "medium": 2, "high": 3}
ALWAYS_MANUAL = {"sonstiges", "unknown", "security_threat", "task"}


def mode(cfg: dict) -> str:
    m = str((cfg.get("auto_approve") or {}).get("mode") or "").lower()
    return m if m in ("shadow", "on") else "off"


def decide(entry: dict, prediction: dict | None, cfg: dict) -> tuple[bool, str]:
    if mode(cfg) == "off":
        return False, ""
    aa = cfg.get("auto_approve") or {}
    allowed = aa.get("categories") or ["advertisement", "newsletter", "application", "booking"]
    min_conf = str(aa.get("min_confidence") or "").lower()
    if min_conf not in ("high", "medium"):
        min_conf = "high"
    cat = str(entry.get("category") or "")
    conf = str((prediction or {}).get("confidence") or "").lower()

    if cfg.get("test_mode"):
        return False, "Test-Modus aktiv"
    if not prediction:
        return False, "keine KI-Einschaetzung"
    if prediction.get("provider") != "langdock" and not aa.get("allow_ollama"):
        return False, "nur Langdock-Ergebnisse werden automatisch freigegeben"
    if entry.get("is_corrected"):
        return False, "korrigierter Eintrag"
    if cat in ALWAYS_MANUAL:  # unabhängig von der Config immer manuell
        return False, "Kategorie wird immer manuell geprueft"
    if cat not in allowed:
        return False, "Kategorie nicht fuer Auto-Freigabe freigeschaltet"
    if entry.get("security_status") != "clean":
        return False, "Sicherheitspruefung nicht bestanden"
    if CONF_RANK.get(conf, 0) < CONF_RANK[min_conf]:
        return False, f"Sicherheit nur '{conf}'"
    if cat == "booking" and (entry.get("action_type") != "forward" or not entry.get("target_email")):
        return False, "Rueckfrage an Gast wird immer manuell geprueft"
    if cat == "newsletter" and not entry.get("unsubscribe_one_click"):
        return False, "Abmeldung ohne One-Click wird manuell geprueft"
    if cat == "application" and not entry.get("target_email"):
        return False, "keine Bewerbungs-Adresse konfiguriert"
    return True, f"Sicherheit '{conf}'"


def apply(entry: dict, prediction: dict | None, cfg: dict) -> None:
    """Schreibt die Entscheidung in den Eintrag. Modus 'on': sofort genehmigen, 'shadow': nur markieren."""
    m = mode(cfg)
    eligible, note = decide(entry, prediction, cfg)
    entry["auto_mode"] = m
    entry["auto_eligible"] = eligible
    entry["auto_note"] = note
    entry["auto_approved"] = False
    if m == "on" and eligible:
        entry["status"] = "approved"
        entry["auto_approved"] = True
