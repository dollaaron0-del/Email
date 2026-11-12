"""Mailzugriff über das lokal installierte Outlook (COM, Paket pywin32).

Wichtig: COM-Objekte gehören zu dem Thread, der sie angelegt hat. Das Backend wird deshalb
im Bot-Thread erzeugt und nur dort benutzt. Das Dashboard öffnet Mails über open_in_outlook(),
das sich eine eigene Verbindung holt.
"""
from __future__ import annotations

from datetime import datetime

from ..log import log
from ..models import Mail
from .base import MailBackend

PR_TRANSPORT_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
OL_FOLDER_INBOX = 6
OL_MAIL_ITEM_CLASS = 43
OL_FLAG_MARKED = 2
OL_IMPORTANCE_HIGH = 2


def _naive(dt) -> datetime:
    """pywin32 liefert Outlook-Zeiten mit Zeitzonen-Etikett, obwohl es lokale Zeiten sind."""
    return datetime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)


def _connect_outlook():
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        app = win32com.client.Dispatch("Outlook.Application")
    return app, app.GetNamespace("MAPI")


def open_in_outlook(mail_id: str) -> None:
    _app, ns = _connect_outlook()
    ns.GetItemFromID(mail_id).Display()


class OutlookBackend(MailBackend):
    name = "outlook"

    def __init__(self) -> None:
        self.app = None
        self.ns = None

    # ── Verbindung ────────────────────────────────────────────────────
    def connect(self) -> None:
        self.app, self.ns = _connect_outlook()

    def ensure_connected(self) -> None:
        try:
            self.ns.GetDefaultFolder(OL_FOLDER_INBOX)
        except Exception:
            log.warning("Outlook-Verbindung verloren, verbinde neu...")
            self.connect()

    @property
    def inbox(self):
        return self.ns.GetDefaultFolder(OL_FOLDER_INBOX)

    def _folder(self, path: str):
        folder = self.inbox
        for part in path.split("/"):
            found = None
            for sub in folder.Folders:
                if sub.Name == part:
                    found = sub
                    break
            folder = found if found is not None else folder.Folders.Add(part)
        return folder

    def _item(self, mail_id: str):
        return self.ns.GetItemFromID(mail_id)

    @staticmethod
    def _prop(item, prop: str) -> str:
        try:
            return str(item.PropertyAccessor.GetProperty(prop) or "")
        except Exception:
            return ""

    def _to_mail(self, item) -> Mail:
        attachments: list[str] = []
        try:
            for i in range(1, item.Attachments.Count + 1):
                attachments.append(str(item.Attachments.Item(i).FileName))
        except Exception:
            pass
        return Mail(
            id=str(item.EntryID),
            message_id=self._prop(item, PR_INTERNET_MESSAGE_ID),
            subject=str(item.Subject or ""),
            body=str(item.Body or ""),
            sender_email=str(item.SenderEmailAddress or ""),
            sender_name=str(item.SenderName or ""),
            received=_naive(item.ReceivedTime),
            is_internal=(str(item.SenderEmailType or "") == "EX"),
            attachments=attachments,
            headers=self._prop(item, PR_TRANSPORT_HEADERS),
        )

    # ── Lesen ─────────────────────────────────────────────────────────
    def list_new(self, since: datetime, limit: int, known_ids: set[str]) -> list[Mail]:
        items = self.inbox.Items
        items.Sort("[ReceivedTime]", True)  # neueste zuerst
        result: list[Mail] = []
        item = items.GetFirst()
        while item is not None:
            try:
                if item.Class == OL_MAIL_ITEM_CLASS:
                    if _naive(item.ReceivedTime) < since:
                        break  # sortiert -> alle weiteren sind älter
                    if str(item.EntryID) not in known_ids:
                        result.append(self._to_mail(item))
                        if len(result) >= limit:
                            break
            except Exception as exc:  # einzelne defekte Elemente überspringen
                log.debug("Element übersprungen: %s", exc)
            item = items.GetNext()
        return result

    def locate(self, mail_id: str, *, folder: str, subject: str, sender_email: str, sender_name: str) -> str | None:
        try:
            if mail_id:
                return str(self._item(mail_id).EntryID)
        except Exception:
            pass
        # Fallback: im Prüf-Ordner nach Betreff + Absender suchen
        try:
            prefix = (subject or "")[:40]
            for m in self._folder(folder).Items:
                try:
                    sender_ok = m.SenderEmailAddress == sender_email or m.SenderName == sender_name
                    if sender_ok and prefix in str(m.Subject or ""):
                        return str(m.EntryID)
                except Exception:
                    continue
        except Exception:
            pass
        return None

    # ── Aktionen ──────────────────────────────────────────────────────
    def move(self, mail_id: str, folder_path: str) -> str:
        moved = self._item(mail_id).Move(self._folder(folder_path))
        return str(moved.EntryID)

    def forward(self, mail_id: str, to: str, subject_prefix: str, body_prefix: str) -> None:
        item = self._item(mail_id)
        fwd = item.Forward()
        fwd.To = to
        fwd.Subject = subject_prefix + str(item.Subject or "")
        fwd.Body = body_prefix + str(fwd.Body or "")
        fwd.Send()

    def reply(self, mail_id: str, text: str) -> None:
        mail = self._item(mail_id).Reply()
        mail.Body = text + "\n\n" + str(mail.Body or "")
        mail.Send()

    def send(self, to: str, subject: str, body: str) -> None:
        mail = self.app.CreateItem(0)
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        mail.Send()

    def flag(self, mail_id: str, important: bool = False) -> None:
        item = self._item(mail_id)
        if important:
            item.Importance = OL_IMPORTANCE_HIGH
        item.FlagStatus = OL_FLAG_MARKED
        item.Save()
