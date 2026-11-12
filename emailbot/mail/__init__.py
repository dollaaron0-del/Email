"""Mailzugriff. Der Rest des Programms spricht nur mit MailBackend - nie direkt mit Outlook."""
from __future__ import annotations

from .base import MailBackend


def make_backend(cfg: dict) -> MailBackend:
    name = str(cfg.get("mail_backend") or "outlook").lower()
    if name == "outlook":
        from .outlook import OutlookBackend
        return OutlookBackend()
    if name == "graph":
        from .graph import GraphBackend
        return GraphBackend(cfg)
    raise ValueError(f"Unbekannter mail_backend '{name}' (erlaubt: outlook, graph)")


def open_in_mail_client(cfg: dict, mail_id: str) -> None:
    """Öffnet eine Mail zur Ansicht (Dashboard-Knopf 'In Outlook')."""
    name = str(cfg.get("mail_backend") or "outlook").lower()
    if name == "outlook":
        from .outlook import open_in_outlook
        open_in_outlook(mail_id)
        return
    raise NotImplementedError("Öffnen im Mailprogramm gibt es nur mit dem Outlook-Backend")


__all__ = ["MailBackend", "make_backend", "open_in_mail_client"]
