"""Queue-Einträge bauen: welche Aktion wird für eine Mail vorgeschlagen?"""
from __future__ import annotations

import uuid
from datetime import datetime
from typing import Callable

from .log import log
from .models import Mail
from .security import unsubscribe_info

DEFAULT_REPLY = ("Sehr geehrte Damen und Herren,\n\nfuer welches Hotel gilt Ihre Anfrage?\n\n{hotels}\n\n"
                 "Mit freundlichen Gruessen")
DEFAULT_REPLY_CORRECTED = ("Sehr geehrte Damen und Herren,\n\nvielen Dank für Ihre Anfrage.\n\n"
                           "Könnten Sie mir bitte mitteilen, in welchem Hotel Sie gerne eine Reservierung "
                           "vornehmen möchten?\n\nIch freue mich auf Ihre Rückmeldung.\n\nMit freundlichen Grüßen")


def hotels_block(cfg: dict) -> str:
    return "\n".join(f"  - {h.get('name')} (Keywords: {', '.join(h.get('keywords') or [])})"
                     for h in cfg.get("hotels") or [])


def find_hotel(name: str, cfg: dict) -> dict | None:
    if not name or name.lower() in ("null", ""):
        return None
    n = name.lower()
    for h in cfg.get("hotels") or []:
        hn = str(h.get("name") or "").lower()
        if hn and (n in hn or hn in n):
            return h
    for h in cfg.get("hotels") or []:
        for kw in h.get("keywords") or []:
            if kw and str(kw).lower() in n:
                return h
    return None


def reply_draft(cfg: dict, language: str = "", detect_language: Callable[[], str] | None = None) -> str:
    templates = cfg.get("reply_templates") or {}
    if language:
        lang = language if templates.get(language) else ("en" if templates.get("en") else "de")
    else:
        lang = detect_language() if detect_language else "de"
    log.info("  -> Sprache erkannt: %s", lang)
    template = templates.get(lang) or templates.get("de") or DEFAULT_REPLY
    hotel_list = "\n".join(f"- {h.get('name')}" for h in cfg.get("hotels") or [])
    return str(template).replace("{hotels}", hotel_list)


def _preview(text: str) -> str:
    b = (text or "").strip()
    return b[:2000] + "…" if len(b) > 2000 else b


def build_queue_entry(mail: Mail, category: str, hotel: str, cfg: dict, *, language: str = "",
                      detect_language: Callable[[], str] | None = None) -> dict:
    entry: dict = {
        "id": str(uuid.uuid4()),
        "entry_id": mail.id,
        "subject": (mail.subject or "")[:80],
        "sender": mail.sender_email,
        "sender_name": mail.sender_name,
        "received": mail.received.strftime("%d.%m.%Y %H:%M"),
        "created": datetime.now().strftime("%d.%m. %H:%M"),
        "category": category,
        "hotel": hotel,
        "status": "pending",
        "proposed_action": "",
        "target_folder": "EmailProcessor/Fehler",
        "target_email": "",
        "unsubscribe_header": "",
        "unsubscribe_http": "",
        "unsubscribe_mailto": "",
        "unsubscribe_one_click": False,
        "body_preview": _preview(mail.body),
    }

    if category == "newsletter":
        u = unsubscribe_info(mail.headers)
        entry.update(unsubscribe_header=u["raw"], unsubscribe_http=u["http"],
                     unsubscribe_mailto=u["mailto"], unsubscribe_one_click=u["one_click"])
        if u["http"]:
            entry["unsubscribe_method"] = "http"
            entry["proposed_action"] = ("RFC 8058 One-Click Abmeldung (automatisch)" if u["one_click"]
                                        else "Abmelde-Webseite aufrufen + Button klicken")
            entry["target_folder"] = "EmailProcessor/Abgemeldet"
        elif u["mailto"]:
            entry["unsubscribe_method"] = "mailto"
            entry["proposed_action"] = "Abmelde-Email senden"
            entry["target_folder"] = "EmailProcessor/Abgemeldet"
        else:
            entry["unsubscribe_method"] = "none"
            entry["proposed_action"] = "Kein Abmelde-Link - manuell abmelden"
            entry["target_folder"] = "EmailProcessor/Newsletter-Manuell"
    elif category == "application":
        entry["proposed_action"] = f"Weiterleiten an: {cfg.get('applications_email')}"
        entry["target_email"] = cfg.get("applications_email") or ""
        entry["target_folder"] = "EmailProcessor/Bewerbung"
    elif category == "advertisement":
        entry["proposed_action"] = "Archivieren - keine weitere Aktion"
        entry["target_folder"] = "EmailProcessor/Werbung"
    elif category == "booking":
        h = find_hotel(hotel, cfg)
        if h:
            entry.update(proposed_action=f"Weiterleiten an {h['name']} ({h['email']})", action_type="forward",
                         target_email=h["email"], target_folder=f"EmailProcessor/Buchung/{h['name']}",
                         hotel=h["name"])
        else:
            log.info("  -> Hotel unklar, generiere Rueckfrage-Entwurf...")
            entry.update(proposed_action="Rueckfrage an Gast senden (Hotel nicht erkennbar)", action_type="reply",
                         reply_draft=reply_draft(cfg, language, detect_language), reply_to=mail.sender_email,
                         target_folder="EmailProcessor/Buchung-Unbekannt")
    elif category == "sonstiges":
        entry["proposed_action"] = "Wichtige Geschäftskorrespondenz - manuell bearbeiten"
        entry["target_folder"] = "EmailProcessor/Sonstiges"
    elif category == "security_threat":
        entry["proposed_action"] = "SICHERHEITSWARNUNG - nicht oeffnen oder ausfuehren!"
        entry["target_folder"] = "EmailProcessor/Sicherheitswarnung"
    else:
        entry["proposed_action"] = "Kategorie unbekannt - manuell prüfen"
        entry["target_folder"] = "EmailProcessor/Unbekannt"
    return entry


def corrected_fields(entry: dict, cfg: dict) -> dict:
    """Felder für einen im Dashboard korrigierten Eintrag (geht danach erneut zur Genehmigung)."""
    cat = entry.get("correct_category")
    f: dict = {"category": cat, "hotel": entry.get("correct_hotel") or "", "action_type": ""}
    if cat == "booking":
        h = find_hotel(entry.get("correct_hotel") or "", cfg) if entry.get("correct_hotel") else None
        if h:
            f.update(target_folder=f"EmailProcessor/Buchung/{h['name']}", target_email=h["email"],
                     proposed_action=f"Weitergeleitet an {h['name']}", action_type="forward", hotel=h["name"])
        else:
            template = (cfg.get("reply_templates") or {}).get("de") or DEFAULT_REPLY_CORRECTED
            f.update(target_folder="EmailProcessor/Buchung-Unbekannt", target_email="", action_type="reply",
                     reply_draft=template, reply_to=entry.get("sender"),
                     proposed_action="Rückfrage an Gast senden (Hotel nicht erkennbar)")
    elif cat == "application":
        f.update(target_folder="EmailProcessor/Bewerbung", target_email=cfg.get("applications_email") or "",
                 proposed_action=f"Weitergeleitet an {cfg.get('applications_email')}")
    elif cat == "advertisement":
        f.update(target_folder="EmailProcessor/Werbung", target_email="", proposed_action="Archiviert als Werbung")
    elif cat == "newsletter":
        f.update(target_folder="EmailProcessor/Newsletter-Manuell", target_email="",
                 proposed_action="Archiviert als Newsletter")
    elif cat == "sonstiges":
        f.update(target_folder="EmailProcessor/Sonstiges", target_email="",
                 proposed_action="Wichtige Geschäftskorrespondenz - manuell bearbeiten")
    else:
        f.update(target_folder="EmailProcessor/Unbekannt", target_email="",
                 proposed_action="Archiviert (Kategorie unbekannt)")
    return f
