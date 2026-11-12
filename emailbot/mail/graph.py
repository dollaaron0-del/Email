"""Platzhalter für den Mailzugriff über Microsoft Graph (für den Betrieb auf einem Server ohne Outlook).

Noch nicht umgesetzt. Voraussetzungen, die die IT schaffen muss:
  - App-Registrierung in Microsoft Entra ID
  - Anwendungsberechtigungen Mail.ReadWrite und Mail.Send (mit Admin-Zustimmung),
    am besten per Application Access Policy auf das eine Postfach beschränkt
  - Client-Secret oder Zertifikat für die App

Abbildung der Schnittstelle auf Graph (Basis: /users/{postfach}/...):
  list_new  -> GET  mailFolders/inbox/messages?$filter=receivedDateTime ge ...&$orderby=receivedDateTime desc
               Felder: id, internetMessageId, subject, body, from, receivedDateTime,
               internetMessageHeaders (für List-Unsubscribe), hasAttachments (+ /attachments für Namen)
  locate    -> GET  messages/{id}  (bei 404: Suche im Prüf-Ordner per $filter)
  move      -> POST messages/{id}/move  {"destinationId": ...}  -> neue id
  forward   -> POST messages/{id}/createForward, Betreff/Text anpassen, dann /send
  reply     -> POST messages/{id}/createReply, Text einfügen, dann /send
  send      -> POST sendMail
  flag      -> PATCH messages/{id}  {"flag": {"flagStatus": "flagged"}, "importance": "high"}
  Ordner    -> GET/POST mailFolders/inbox/childFolders

Die Internet-Message-ID in den Trainingsdaten ist in Graph als internetMessageId verfügbar.
"""
from __future__ import annotations

from .base import MailBackend


class GraphBackend(MailBackend):  # pragma: no cover - noch nicht umgesetzt
    name = "graph"

    def __init__(self, cfg: dict) -> None:
        raise NotImplementedError(
            "Microsoft Graph ist vorbereitet, aber noch nicht umgesetzt. "
            "Bis dahin in config.json \"mail_backend\": \"outlook\" verwenden."
        )

    def connect(self): ...
    def list_new(self, since, limit, known_ids): ...
    def locate(self, mail_id, *, folder, subject, sender_email, sender_name): ...
    def move(self, mail_id, folder_path): ...
    def forward(self, mail_id, to, subject_prefix, body_prefix): ...
    def reply(self, mail_id, text): ...
    def send(self, to, subject, body): ...
    def flag(self, mail_id, important=False): ...
