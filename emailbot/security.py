"""Regelbasierte Prüfungen: gefährliche Anhänge/Links, Spam-Absender, Kollegen-Mails, Abmelde-Header."""
from __future__ import annotations

import os
import re

from .models import Mail

DANGER_EXTS = {
    ".exe", ".bat", ".cmd", ".com", ".pif", ".scr", ".vbs", ".vbe", ".js", ".jse", ".wsf", ".wsh",
    ".ps1", ".ps2", ".msi", ".msp", ".jar", ".reg", ".hta", ".lnk", ".iso", ".img",
}
WARN_EXTS = {".zip", ".rar", ".7z", ".gz", ".tar", ".docm", ".xlsm", ".pptm", ".dotm", ".xltm", ".xlam"}

URL_RE = re.compile(r'https?://[^\s<>"]+', re.I)
IP_URL_RE = re.compile(r"https?://\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", re.I)
PHISH_URL_RE = re.compile(
    r"(verify.{0,10}account|confirm.{0,10}password|update.{0,10}billing|paypal.{0,20}\.(?!com)"
    r"|microsoft.{0,20}\.(?!com)|apple.{0,20}\.(?!com)|secure.{0,10}login\.)",
    re.I,
)
WEBMAIL_RE = re.compile(r"@(gmail|hotmail|yahoo|outlook|web|gmx|aol|icloud)\.(com|de|at|ch)$", re.I)


def check_rules(mail: Mail) -> dict:
    """Anhänge und Links prüfen. Ergebnis: threat_level none/warn/danger + Gründe."""
    reasons: list[str] = []
    level = "none"

    for name in mail.attachments:
        ext = os.path.splitext(name)[1].lower()
        if ext in DANGER_EXTS:
            reasons.append(f"Gefaehrlicher Anhang: {name}")
            level = "danger"
        elif ext in WARN_EXTS:
            reasons.append(f"Verdaechtiger Anhang (Makro/Archiv): {name}")
            if level == "none":
                level = "warn"

    body = mail.body or ""
    if IP_URL_RE.search(body):
        reasons.append("Link zu IP-Adresse (kein Domainname)")
        if level == "none":
            level = "warn"
    for url in URL_RE.findall(body):
        if len(url) > 200:
            reasons.append(f"Sehr langer/verdaechtiger Link ({url[:60]}...)")
            if level == "none":
                level = "warn"
            break
        if PHISH_URL_RE.search(url):
            reasons.append(f"Phishing-Muster in URL: {url[:80]}")
            level = "danger"

    return {"threat_level": level, "reasons": reasons, "is_threat": level != "none"}


def merge_llm_phishing(sec: dict, is_phishing: bool, confidence: str, reason: str) -> dict:
    """Übernimmt eine KI-Phishing-Einschätzung (Langdock oder Ollama) in das Prüfergebnis."""
    if is_phishing:
        sec["reasons"].append(f"LLM-Phishing-Erkennung ({confidence}): {reason}")
        if confidence in ("high", "medium"):
            sec["threat_level"] = "danger"
        elif sec["threat_level"] == "none":
            sec["threat_level"] = "warn"
        sec["is_threat"] = sec["threat_level"] != "none"
    return sec


def is_suspicious_spam(mail: Mail) -> bool:
    """Webmail-Absender ohne echten Namen oder mit zufällig wirkender Adresse."""
    email = (mail.sender_email or "").lower()
    name = mail.sender_name or ""
    if not WEBMAIL_RE.search(email):
        return False
    local = email.split("@", 1)[0]
    if name == mail.sender_email or name.lower() == local:
        return True
    if len(local) > 12 and not re.search(r"[._]", local) and re.search(r"\d", local) \
            and not re.match(r"^\w+[12][0-9]{3}$", local):
        return True
    return False


def should_skip(mail: Mail, cfg: dict) -> bool:
    """Kollegen-Mails (Exchange-intern, eigene Domains, einzelne Adressen) nicht klassifizieren."""
    if mail.is_internal:
        return True
    sender = (mail.sender_email or "").lower()
    for domain in cfg.get("skip_domains") or []:
        if domain and sender.endswith("@" + str(domain).lower()):
            return True
    for addr in cfg.get("skip_senders") or []:
        if addr and sender == str(addr).lower():
            return True
    return False


def unsubscribe_info(headers: str) -> dict:
    """List-Unsubscribe-Header auswerten (auch über mehrere Zeilen umbrochene Header)."""
    info = {"raw": "", "mailto": "", "http": "", "one_click": False}
    if not headers:
        return info
    unfolded = re.sub(r"\r?\n[ \t]+", " ", headers)
    m = re.search(r"(?im)^List-Unsubscribe:[ \t]*(.+?)[ \t\r]*$", unfolded)
    if m:
        raw = m.group(1).strip()
        info["raw"] = raw
        mm = re.search(r"<mailto:([^>]+)>", raw, re.I)
        if mm:
            info["mailto"] = mm.group(1)
        mh = re.search(r"<(https?://[^>]+)>", raw, re.I)
        if mh:
            info["http"] = mh.group(1)
    if re.search(r"(?im)^List-Unsubscribe-Post:[ \t]*List-Unsubscribe=One-Click", unfolded):
        info["one_click"] = True
    return info
