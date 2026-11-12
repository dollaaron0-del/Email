"""Datenmodell einer Mail - unabhängig davon, ob sie aus Outlook oder Microsoft Graph kommt."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class Mail:
    id: str                    # Kennung im Postfach (Outlook: EntryID, Graph: message id) - ändert sich beim Verschieben
    message_id: str            # Internet-Message-ID - bleibt gleich, auch in Microsoft Graph
    subject: str
    body: str
    sender_email: str
    sender_name: str
    received: datetime         # lokale Zeit, ohne Zeitzone
    is_internal: bool = False  # Absender im eigenen Exchange (Kollegen)
    attachments: list[str] = field(default_factory=list)
    headers: str = ""          # Transport-Header (für List-Unsubscribe)
