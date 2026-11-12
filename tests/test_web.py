"""Dashboard-API - gleiche Adressen und Antworten wie dashboard_server.ps1."""
import codecs
import json

import pytest

from emailbot.web import create_app


class FakeRunner:
    def __init__(self):
        self.running = True

    def pause(self):
        self.running = False

    def resume(self):
        self.running = True


@pytest.fixture
def client(base_dir):
    runner = FakeRunner()
    opened = []
    app = create_app(base_dir, runner, open_mail=lambda cfg, mail_id: opened.append(mail_id))
    app.testing = True
    c = app.test_client()
    c.runner, c.opened, c.base_dir = runner, opened, base_dir
    return c


def write_queue_ps_style(base_dir, entries):
    """So schreibt die PowerShell-Version: UTF-8 mit BOM."""
    text = "[\n" + ",\n".join(json.dumps(e, ensure_ascii=False) for e in entries) + "\n]"
    (base_dir / "queue.json").write_bytes(codecs.BOM_UTF8 + text.encode("utf-8"))


def entry(i, status="pending", received="12.11.2026 10:00", **kw):
    return {"id": i, "entry_id": "e" + i, "subject": f"Mail {i}", "status": status, "received": received,
            "category": "booking", **kw}


def test_stats_fields(client):
    d = client.get("/api/stats").get_json()
    assert d["bot_running"] is True and d["llm_provider"] == "langdock" and d["auto_mode"] == "on"
    assert d["langdock_key_set"] is True and d["langdock_key_name"] == "LANGDOCK_API_KEY"
    assert d["training_enabled"] is True and d["ollama_ok"] is True and d["test_mode"] is False
    assert "testkey" not in json.dumps(d)


def test_queue_filter_sort_and_ps_compat(client):
    write_queue_ps_style(client.base_dir, [
        entry("a", received="11.11.2026 09:00"),
        {"value": [entry("b", received="12.11.2026 09:00")], "Count": 1},  # alte verschachtelte Struktur
        entry("c", status="done"),
        entry("d", status="error"),
        {"id": "t", "category": "task", "status": "open_task", "received": "12.11.2026 08:00"},
    ])
    assert [e["id"] for e in client.get("/api/queue").get_json()] == ["b", "a"]
    assert [e["id"] for e in client.get("/api/errors").get_json()] == ["d"]
    assert [e["id"] for e in client.get("/api/tasks").get_json()] == ["t"]


def test_approve_reject_ignore(client):
    write_queue_ps_style(client.base_dir, [entry("a"), entry("b"), entry("c"), entry("d", status="error")])
    assert client.post("/api/queue/approve", json={"id": "a", "reply_text": "Hallo"}).get_json() == {"success": True}
    client.post("/api/queue/reject", json={"id": "b", "correct_category": "newsletter", "rejection_reason": "falsch"})
    client.post("/api/queue/ignore", json={"id": "c"})
    client.post("/api/errors/retry", json={"id": "d"})
    q = {e["id"]: e for e in json.loads((client.base_dir / "queue.json").read_text(encoding="utf-8"))}
    assert q["a"]["status"] == "approved" and q["a"]["reply_text"] == "Hallo"
    assert q["b"]["status"] == "rejected" and q["b"]["correct_category"] == "newsletter"
    assert q["c"]["status"] == "ignored" and q["d"]["status"] == "approved"


def test_config_roundtrip_keeps_unknown_keys(client):
    cfg = client.get("/api/config").get_json()
    cfg["auto_approve"]["mode"] = "shadow"
    cfg["eigenes_feld"] = "bleibt"
    assert client.post("/api/config", data=json.dumps(cfg)).get_json() == {"success": True}
    saved = json.loads((client.base_dir / "config.json").read_text(encoding="utf-8"))
    assert saved["auto_approve"]["mode"] == "shadow" and saved["eigenes_feld"] == "bleibt"


def test_config_rejects_invalid_json(client):
    before = (client.base_dir / "config.json").read_text(encoding="utf-8")
    resp = client.post("/api/config", data="{kaputt")
    assert resp.status_code == 500 and "error" in resp.get_json()
    assert (client.base_dir / "config.json").read_text(encoding="utf-8") == before


def test_bot_start_stop(client):
    client.post("/api/bot/stop")
    assert client.get("/api/stats").get_json()["bot_running"] is False
    client.post("/api/bot/start")
    assert client.get("/api/stats").get_json()["bot_running"] is True


def test_open_email_and_logs(client):
    assert client.post("/api/email/open", json={"entry_id": "abc"}).get_json() == {"success": True}
    assert client.opened == ["abc"]
    (client.base_dir / "email_processor.log").write_text("\n".join(f"Zeile {i}" for i in range(60)), encoding="utf-8")
    logs = client.get("/api/logs").get_json()
    assert len(logs) == 40 and logs[-1] == "Zeile 59"


def test_static_files_and_traversal(client):
    assert b"Dashboard" in client.get("/").data
    assert client.get("/../config.json").status_code == 404
    assert client.get("/fehlt.html").status_code == 404
