"""Schnittstelle für den Mailzugriff.

Ordnerpfade sind immer relativ zum Posteingang, z. B. "EmailProcessor/Buchung/Hotel X".
Fehlende Ordner werden angelegt.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from datetime import datetime

from ..models import Mail


class MailBackend(ABC):
    name = "base"

    @abstractmethod
    def connect(self) -> None:
        """Verbindung herstellen (wird im Bot-Thread aufgerufen)."""

    def ensure_connected(self) -> None:
        """Verbindung prüfen und bei Bedarf neu aufbauen."""

    @abstractmethod
    def list_new(self, since: datetime, limit: int, known_ids: set[str]) -> list[Mail]:
        """Neueste Mails im Posteingang ab 'since', ohne bekannte IDs, höchstens 'limit' Stück."""

    @abstractmethod
    def locate(self, mail_id: str, *, folder: str, subject: str, sender_email: str, sender_name: str) -> str | None:
        """Gibt eine gültige ID zurück. Ist mail_id ungültig, wird im Ordner nach Betreff + Absender gesucht."""

    @abstractmethod
    def move(self, mail_id: str, folder_path: str) -> str:
        """Verschiebt die Mail und gibt die neue ID zurück."""

    @abstractmethod
    def forward(self, mail_id: str, to: str, subject_prefix: str, body_prefix: str) -> None:
        """Leitet weiter. Betreff = subject_prefix + Originalbetreff, Text = body_prefix + Originaltext."""

    @abstractmethod
    def reply(self, mail_id: str, text: str) -> None:
        """Antwortet dem Absender (text wird über dem zitierten Original eingefügt)."""

    @abstractmethod
    def send(self, to: str, subject: str, body: str) -> None:
        """Schickt eine neue Mail (z. B. Abmeldung per mailto)."""

    @abstractmethod
    def flag(self, mail_id: str, important: bool = False) -> None:
        """Markiert die Mail (Fähnchen), optional mit hoher Wichtigkeit."""
