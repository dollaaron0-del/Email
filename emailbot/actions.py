"""Aktionen nach Genehmigung ausführen: weiterleiten, antworten, abmelden, verschieben."""
from __future__ import annotations

import re
from urllib.parse import unquote, urljoin

import requests

from .log import log
from .mail.base import MailBackend

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/120.0.0.0 Safari/537.36")
HEADERS = {"User-Agent": UA, "Accept": "text/html,application/xhtml+xml,*/*;q=0.9",
           "Accept-Language": "de-DE,de;q=0.9,en;q=0.8"}

# Phrasen für erkannte Abmeldung (DE/EN/FR/NL/ES)
CONFIRM_RE = re.compile(
    r"(erfolgreich.{0,12}abgemeldet|abmeldung.erfolgreich|erfolgreich.entfernt|abonnement.gek.ndigt|unsubscribed"
    r"|successfully.removed|opt.?out.success|you.ve been unsubscribed|you have been unsubscribed|you.re.unsubscribed"
    r"|have been removed|no longer receive|désabonné|vous avez .t. désabonné|afgemeld|avmeldt|dado de baja)",
    re.I,
)
FORM_RE = re.compile(r'<form[^>]*?action="?([^\s">]+)"?[^>]*?>(.*?)</form>', re.I | re.S)
UNSUB_WORDS_RE = re.compile(r"(unsubscribe|abmeld|opt.?out|austragen|confirm|remove|abbestell)", re.I)
INPUT_RE = re.compile(r"<input([^>]+)>", re.I)
BUTTON_RE = re.compile(r'<button[^>]+name="([^"]+)"[^>]*(?:value="([^"]*)")?', re.I)
LINK_RE = re.compile(r'<a[^>]+href="(https?://[^"]+)"[^>]*>[^<]{0,80}(?:unsubscribe|abmeld|opt.?out|austragen|abbestell)'
                     r"[^<]{0,40}</a>", re.I)


def web_unsubscribe(url: str, one_click: bool = False, http=None) -> dict:
    s = http or requests.Session()
    if one_click:  # RFC 8058: einfacher POST, keine Seiteninteraktion nötig
        try:
            r = s.post(url, data="List-Unsubscribe=One-Click", headers={
                **HEADERS, "Content-Type": "application/x-www-form-urlencoded"}, timeout=15)
            if r.status_code < 400:
                return {"method": "rfc8058", "message": f"RFC 8058 One-Click Abmeldung erfolgreich (HTTP {r.status_code})"}
        except requests.RequestException:
            log.warning("  -> RFC 8058 POST fehlgeschlagen, versuche normales GET...")
    try:
        resp = s.get(url, headers=HEADERS, timeout=20)
        html = resp.text
        if CONFIRM_RE.search(html):
            return {"method": "get", "message": "Seite bestätigt Abmeldung direkt (kein Klick nötig)"}

        for action, content in FORM_RE.findall(html):
            if not UNSUB_WORDS_RE.search(content):
                continue
            fields: dict[str, str] = {}
            for inp in INPUT_RE.findall(content):
                itype = (re.search(r'type="?(\w+)', inp, re.I) or [None, "text"])[1]
                name = re.search(r'name="([^">\s]+)"', inp, re.I)
                value = re.search(r'value="([^"]*)"', inp, re.I)
                if name and itype.lower() in ("hidden", "submit", "email", "text", "checkbox"):
                    fields[name.group(1)] = value.group(1) if value else ""
            for bname, bval in BUTTON_RE.findall(content):
                fields[bname] = bval
            post = s.post(urljoin(resp.url, action), data=fields, headers=HEADERS, timeout=20)
            if CONFIRM_RE.search(post.text):
                return {"method": "post_form", "message": "Formular abgeschickt - Seite bestätigt Abmeldung"}
            return {"method": "post_form", "message": "Formular abgeschickt (keine explizite Bestätigung - bitte prüfen)"}

        m = LINK_RE.search(html)
        if m:
            try:
                link = s.get(m.group(1), headers=HEADERS, timeout=15)
                if CONFIRM_RE.search(link.text):
                    return {"method": "link_click", "message": "Abmelde-Link geklickt - Seite bestätigt Abmeldung"}
                return {"method": "link_click", "message": "Abmelde-Link geklickt (Bestätigung unklar - bitte prüfen)"}
            except requests.RequestException as exc:
                log.warning("  -> Abmelde-Link fehlgeschlagen: %s", exc)
        return {"method": "get", "message": "Seite aufgerufen, kein Abmelde-Formular/Link gefunden - bitte manuell prüfen"}
    except requests.RequestException as exc:
        return {"method": "error", "message": f"Fehler beim Aufrufen der Abmelde-Seite: {exc}"}


def execute_entry(entry: dict, backend: MailBackend, http=None) -> str:
    """Führt die vorgeschlagene Aktion aus. Rückgabe: 'done' oder 'error'."""
    mail_id = backend.locate(entry.get("entry_id") or "", folder="EmailProcessor/Zur Pruefung",
                             subject=entry.get("subject") or "", sender_email=entry.get("sender") or "",
                             sender_name=entry.get("sender_name") or "")
    if not mail_id:
        log.warning("  -> E-Mail nicht mehr auffindbar (EntryID ungültig)")
        return "error"

    try:
        cat = entry.get("category")
        if cat == "newsletter":
            http_url = entry.get("unsubscribe_http") or ""
            mailto = entry.get("unsubscribe_mailto") or ""
            header = entry.get("unsubscribe_header") or ""
            if not http_url and not mailto and header:  # ältere Einträge
                mm = re.search(r"<mailto:([^>]+)>", header)
                mh = re.search(r"<(https?://[^>]+)>", header)
                mailto = mm.group(1) if mm else ""
                http_url = mh.group(1) if mh else ""
            if http_url:
                result = web_unsubscribe(http_url, bool(entry.get("unsubscribe_one_click")), http)
                log.info("  -> Web-Abmeldung (%s): %s", result["method"], result["message"])
            elif mailto:
                addr, _, query = mailto.partition("?")
                subject = "unsubscribe"
                m = re.search(r"subject=([^&]+)", query, re.I)
                if m:
                    subject = unquote(m.group(1))
                backend.send(addr, subject, "unsubscribe")
                log.info("  -> Abmelde-Email gesendet an %s (Betreff: %s)", addr, subject)
        elif cat == "application":
            backend.forward(mail_id, entry.get("target_email") or "", "BEWERBUNG: ",
                            f"Weitergeleitet von: {entry.get('sender')}\n{'-' * 60}\n")
        elif cat == "booking":
            if entry.get("action_type") == "reply":
                text = entry.get("reply_text") or entry.get("reply_draft") or ""
                backend.reply(mail_id, text)
                log.info("  -> Rueckfrage gesendet an %s", entry.get("reply_to") or entry.get("sender"))
            elif entry.get("target_email"):
                backend.forward(mail_id, entry["target_email"], f"BUCHUNGSANFRAGE - {entry.get('hotel')}: ",
                                f"Von: {entry.get('sender')}\n{'-' * 60}\n")
        elif cat == "sonstiges":
            try:  # rot markieren + hohe Wichtigkeit, damit die Mail sofort auffällt
                backend.flag(mail_id, important=True)
            except Exception:
                pass
        elif cat == "security_threat":
            log.info("  -> Sicherheitswarnung bestaetigt, bleibt in %s", entry.get("target_folder"))
            return "done"
        # Werbung und andere: nur verschieben
        backend.move(mail_id, entry.get("target_folder") or "EmailProcessor/Unbekannt")
        log.info("  -> Ausgefuehrt: %s", entry.get("proposed_action"))
        return "done"
    except Exception as exc:
        log.error("  -> Fehler beim Ausfuehren: %s", exc)
        return "error"
