"""Test-Hilfen: Postfach im Speicher statt Outlook, lokaler Testserver statt Langdock/Ollama."""
from __future__ import annotations

import json
import threading
import uuid
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest

from emailbot.llm import langdock
from emailbot.mail.base import MailBackend
from emailbot.models import Mail

ALPHA_REASON = "Buchungsanfrage für Hotel Alpha – München"


# ── Postfach-Attrappe ─────────────────────────────────────────────────
class FakeBackend(MailBackend):
    name = "fake"

    def __init__(self) -> None:
        self.items: dict[str, Mail] = {}
        self.folder_of: dict[str, str] = {}
        self.moves: list[tuple[str, str]] = []
        self.sent: list[tuple] = []
        self.flags: list[tuple[str, bool]] = []

    def connect(self) -> None:
        pass

    def add(self, mail: Mail) -> Mail:
        self.items[mail.id] = mail
        self.folder_of[mail.id] = "Posteingang"
        return mail

    def list_new(self, since, limit, known_ids):
        inbox = [m for i, m in self.items.items() if self.folder_of[i] == "Posteingang"]
        inbox.sort(key=lambda m: m.received, reverse=True)
        return [m for m in inbox if m.received >= since and m.id not in known_ids][:limit]

    def locate(self, mail_id, *, folder, subject, sender_email, sender_name):
        return mail_id if mail_id in self.items else None

    def move(self, mail_id, folder_path):
        mail = self.items.pop(mail_id)  # KeyError wie "nicht gefunden"
        self.folder_of.pop(mail_id)
        new_id = "moved-" + uuid.uuid4().hex
        self.items[new_id] = mail
        self.folder_of[new_id] = folder_path
        self.moves.append((mail.subject, folder_path))
        return new_id

    def forward(self, mail_id, to, subject_prefix, body_prefix):
        self.sent.append(("forward", to, subject_prefix + self.items[mail_id].subject, body_prefix))

    def reply(self, mail_id, text):
        self.sent.append(("reply", self.items[mail_id].sender_email, text))

    def send(self, to, subject, body):
        self.sent.append(("send", to, subject, body))

    def flag(self, mail_id, important=False):
        self.flags.append((mail_id, important))


def make_mail(subject: str, body: str = "Text", sender: str = "max.muster@gmail.com", name: str = "Max Muster",
              headers: str = "", attachments: list[str] | None = None, internal: bool = False) -> Mail:
    return Mail(id=uuid.uuid4().hex, message_id=f"<{uuid.uuid4().hex}@example.com>", subject=subject, body=body,
                sender_email=sender, sender_name=name, received=datetime.now(), is_internal=internal,
                attachments=attachments or [], headers=headers)


# ── Testserver für Langdock, Ollama und Abmelde-Seiten ────────────────
def _answer(body: str) -> dict:
    a = {"category": "advertisement", "hotel_name": None, "confidence": "medium", "reason": "Werbung",
         "language": "de", "is_phishing": False, "phishing_confidence": "low", "phishing_reason": ""}
    if "MOCK_BOOKING_ALPHA" in body:
        a.update(category="booking", hotel_name="Hotel Alpha", confidence="high", reason=ALPHA_REASON)
    elif "MOCK_NEWSLETTER" in body:
        a.update(category="newsletter", confidence="high", reason="Newsletter")
    elif "MOCK_SONSTIGES" in body:
        a.update(category="sonstiges", confidence="high", reason="Rechnung")
    elif "MOCK_BOOKING_UNKNOWN" in body:
        a.update(category="booking", confidence="high", reason="Hotel unklar", language="en")
    elif "MOCK_PHISH" in body:
        a.update(is_phishing=True, phishing_confidence="high", phishing_reason="Gefälschte Bank")
    elif "MOCK_MEDIUM" in body:
        a.update(category="application", confidence="medium", reason="vielleicht Bewerbung")
    return a


class MockHandler(BaseHTTPRequestHandler):
    def log_message(self, *args):  # ruhig bleiben
        pass

    def _send(self, code: int, payload, content_type: str = "application/json") -> None:
        raw = payload if isinstance(payload, bytes) else (
            payload if isinstance(payload, str) else json.dumps(payload, ensure_ascii=False)).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)  # bewusst ohne charset
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _handle(self, method: str) -> None:
        length = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(length).decode("utf-8") if length else ""
        path = self.path.split("?")[0]
        self.server.requests.append({"path": path, "method": method, "auth": self.headers.get("Authorization"),
                                     "body": body})
        needs_auth = path.startswith("/openai/") or path.startswith("/agent/")
        if needs_auth and self.headers.get("Authorization") != "Bearer testkey":
            return self._send(401, {"error": "invalid api key"})
        if path.endswith("/v1/models"):
            return self._send(200, {"data": [{"id": "test-model"}, {"id": "other-model"}]})
        if path.startswith("/openai/") and path.endswith("/chat/completions"):
            content = "Hier das Ergebnis:\n```json\n" + json.dumps(_answer(body), ensure_ascii=False) + "\n```"
            return self._send(200, {"choices": [{"index": 0, "message": {"role": "assistant", "content": content}}]})
        if path == "/agent/v1/chat/completions":
            ans = _answer(body)
            if '"output"' in body:
                return self._send(200, {"messages": [{"id": "m1", "role": "assistant", "content": "erledigt"}],
                                        "output": ans})
            return self._send(200, {"messages": [{"id": "m1", "role": "assistant", "content": json.dumps(ans)}]})
        if path == "/api/chat":
            if "Phishing" in body:
                c = '{"is_phishing": false, "confidence": "low", "reason": ""}'
            elif "Detect the language" in body:
                c = "de"
            else:
                c = '{"category":"advertisement","hotel_name":null,"confidence":"high","reason":"Ollama sagt Werbung"}'
            return self._send(200, {"message": {"role": "assistant", "content": c}}, "application/json; charset=utf-8")
        if path == "/api/tags":
            return self._send(200, {"models": []})
        return self._send(200, "Sie wurden erfolgreich abgemeldet", "text/plain")

    def do_GET(self):
        self._handle("GET")

    def do_POST(self):
        self._handle("POST")


@pytest.fixture(scope="session")
def mock_server():
    server = ThreadingHTTPServer(("127.0.0.1", 0), MockHandler)
    server.requests = []
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    yield server
    server.shutdown()


@pytest.fixture(autouse=True)
def api_key(monkeypatch):
    """Nie den echten Langdock-Key dieses Rechners verwenden."""
    monkeypatch.setattr(langdock, "_read_registry_env", lambda name: None)
    monkeypatch.setenv("LANGDOCK_API_KEY", "testkey")


@pytest.fixture
def base_cfg(mock_server) -> dict:
    url = f"http://127.0.0.1:{mock_server.server_address[1]}"
    return {
        "hotels": [{"name": "Hotel Alpha", "email": "alpha@example.com", "keywords": ["alpha"]},
                   {"name": "Hotel Beta", "email": "beta@example.com", "keywords": ["beta"]}],
        "applications_email": "hr@example.com",
        "ollama_model": "llama3.2", "ollama_url": f"{url}/api/chat",
        "check_interval_minutes": 5, "scan_max_age_days": 2,
        "skip_domains": ["meine-firma.de"], "skip_senders": [],
        "llm_provider": "langdock",
        "langdock": {"model": "test-model", "region": "eu", "base_url": url, "fallback_to_ollama": True},
        "auto_approve": {"mode": "on", "min_confidence": "high",
                         "categories": ["advertisement", "newsletter", "application", "booking"]},
        "training_data": {"enabled": True, "dir": "training_data", "body_chars": 4000},
        "reply_templates": {"de": "Hallo, welches Hotel? {hotels}", "en": "Hello, which hotel? {hotels}"},
    }


@pytest.fixture
def base_dir(tmp_path, base_cfg) -> Path:
    (tmp_path / "config.json").write_text(json.dumps(base_cfg), encoding="utf-8")
    (tmp_path / "www").mkdir()
    (tmp_path / "www" / "index.html").write_text("<html>Dashboard</html>", encoding="utf-8")
    return tmp_path
