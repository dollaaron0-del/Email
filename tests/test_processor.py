"""Komplettablauf mit Postfach-Attrappe: scannen, einordnen, automatisch/manuell ausführen, Trainingsdaten."""
import json

import pytest

from emailbot.config import load_config
from emailbot.export_training import merge_events
from emailbot.processor import Processor

from .conftest import ALPHA_REASON, FakeBackend, make_mail


def by_subject(proc, part):
    return next(e for e in proc.queue.load() if part in e["subject"])


def events(base_dir):
    files = list((base_dir / "training_data").glob("emails_*.jsonl"))
    assert len(files) == 1
    return [json.loads(line) for line in files[0].read_text(encoding="utf-8").splitlines()]


def requests_since(server, start):
    return server.requests[start:]


@pytest.fixture
def scenario_a(base_dir, mock_server):
    """Langdock, Auto-Freigabe AN, Trainingsdaten AN - acht typische Mails."""
    backend = FakeBackend()
    port = mock_server.server_address[1]
    nl_header = (f"List-Unsubscribe: <http://127.0.0.1:{port}/unsub>\r\n"
                 "List-Unsubscribe-Post: List-Unsubscribe=One-Click\r\n")
    for m in [
        make_mail("Zimmer MOCK_BOOKING_ALPHA", "Wir möchten ein Zimmer im Alpha buchen.", "anna.schmidt@gmail.com",
                  "Anna Schmidt"),
        make_mail("News MOCK_NEWSLETTER", "Unser Newsletter", "news@shop.example.com", "Shop News",
                  headers=nl_header),
        make_mail("Rechnung MOCK_SONSTIGES", "Anbei die Rechnung", "buchhaltung@lieferant.example.com", "Lieferant"),
        make_mail("Room MOCK_BOOKING_UNKNOWN", "We would like to book a room.", "john.doe@example.com", "John Doe"),
        make_mail("Konto MOCK_PHISH", "Bitte bestaetigen Sie Ihr Konto", "service@bank-example.net", "Bank"),
        make_mail("Rechnung exe", "Siehe Anhang", "x@example.org", "X Y", attachments=["rechnung.exe"]),
        make_mail("Frage Kollege", "Hallo", "kollege@meine-firma.de", "Kollege"),
        make_mail("Bewerbung MOCK_MEDIUM", "Ich bewerbe mich vielleicht", "bewerber@example.com", "Bewerber"),
    ]:
        backend.add(m)
    proc = Processor(base_dir, backend)
    cfg = load_config(base_dir)
    start = len(mock_server.requests)
    proc.scan_inbox(cfg)
    proc.process_queue(cfg)
    return proc, backend, cfg, requests_since(mock_server, start)


def test_scenario_a_requests(scenario_a):
    proc, backend, cfg, reqs = scenario_a
    assert len(proc.queue.load()) == 8
    assert len([r for r in reqs if r["path"].startswith("/openai/")]) == 6  # nicht für .exe und Kollegen-Mail
    assert not [r for r in reqs if r["path"] == "/api/chat"]  # Ollama nicht benutzt


def test_scenario_a_booking_auto_forwarded(scenario_a):
    proc, backend, cfg, _ = scenario_a
    bk = by_subject(proc, "MOCK_BOOKING_ALPHA")
    assert bk["status"] == "done" and bk["auto_approved"] is True and bk["hotel"] == "Hotel Alpha"
    assert bk["llm_provider"] == "langdock" and bk["message_id"].startswith("<")
    assert ("forward", "alpha@example.com", "BUCHUNGSANFRAGE - Hotel Alpha: Zimmer MOCK_BOOKING_ALPHA",
            "Von: anna.schmidt@gmail.com\n" + "-" * 60 + "\n") in backend.sent
    assert ("Zimmer MOCK_BOOKING_ALPHA", "EmailProcessor/Buchung/Hotel Alpha") in backend.moves


def test_scenario_a_newsletter_one_click(scenario_a):
    proc, backend, cfg, reqs = scenario_a
    assert by_subject(proc, "MOCK_NEWSLETTER")["status"] == "done"
    unsub = [r for r in reqs if r["path"] == "/unsub"]
    assert len(unsub) == 1 and unsub[0]["method"] == "POST" and unsub[0]["body"] == "List-Unsubscribe=One-Click"
    assert ("News MOCK_NEWSLETTER", "EmailProcessor/Abgemeldet") in backend.moves


def test_scenario_a_manual_cases(scenario_a):
    proc, backend, cfg, _ = scenario_a
    so = by_subject(proc, "MOCK_SONSTIGES")
    assert so["status"] == "pending" and "immer manuell" in so["auto_note"]
    bu = by_subject(proc, "MOCK_BOOKING_UNKNOWN")
    assert bu["status"] == "pending" and bu["action_type"] == "reply"
    assert bu["reply_draft"].startswith("Hello, which hotel?") and "Rueckfrage" in bu["auto_note"]
    med = by_subject(proc, "MOCK_MEDIUM")
    assert med["status"] == "pending" and "medium" in med["auto_note"]
    assert not [s for s in backend.sent if s[0] == "reply"]  # nie automatisch an Gäste schreiben


def test_scenario_a_security_and_colleague(scenario_a):
    proc, backend, cfg, _ = scenario_a
    ph = by_subject(proc, "MOCK_PHISH")
    assert ph["category"] == "security_threat" and "LLM-Phishing" in ph["security_warning"]
    exe = by_subject(proc, "Rechnung exe")
    assert exe["category"] == "security_threat" and "rechnung.exe" in exe["security_warning"]
    assert ("Konto MOCK_PHISH", "EmailProcessor/Sicherheitswarnung") in backend.moves
    kol = by_subject(proc, "Frage Kollege")
    assert kol["category"] == "task" and kol["status"] == "open_task"
    assert any(not important for _, important in backend.flags)


def test_scenario_a_stats_and_training(scenario_a, base_dir):
    proc, backend, cfg, _ = scenario_a
    stats = json.loads((base_dir / "stats.json").read_text(encoding="utf-8"))
    assert len([r for r in stats["recent"] if r["auto"]]) == 2
    evs = events(base_dir)
    preds = [e for e in evs if e["event"] == "prediction"]
    decs = [e for e in evs if e["event"] == "decision"]
    assert len(preds) == 7  # alle außer Kollegen-Mail
    assert len([d for d in decs if d["decision"] == "auto_approved"]) == 2
    bk = by_subject(proc, "MOCK_BOOKING_ALPHA")
    p = next(e for e in preds if e["id"] == bk["id"])
    assert p["input"]["sender_domain"] == "gmail.com" and p["input"]["body"].startswith("Wir möchten")
    assert p["prediction"]["reason"] == ALPHA_REASON and p["prompt_version"] == "langdock-v1"
    nl = next(e for e in preds if e["id"] == by_subject(proc, "MOCK_NEWSLETTER")["id"])
    assert nl["input"]["has_list_unsubscribe"] and nl["input"]["one_click_unsubscribe"]
    exe = next(e for e in preds if e["id"] == by_subject(proc, "Rechnung exe")["id"])
    assert exe["input"]["attachments"] == ["rechnung.exe"] and exe["extra"]["threat_level"] == "danger"
    raw = (base_dir / "training_data").glob("emails_*.jsonl").__next__().read_text(encoding="utf-8")
    assert "anna.schmidt@" not in raw  # keine vollständige Absenderadresse


def test_scenario_a_dashboard_decisions_and_export(scenario_a, base_dir):
    proc, backend, cfg, _ = scenario_a
    so, bu, ph, med = (by_subject(proc, s)["id"] for s in
                       ("MOCK_SONSTIGES", "MOCK_BOOKING_UNKNOWN", "MOCK_PHISH", "MOCK_MEDIUM"))
    proc.queue.update(so, {"status": "approved"})
    proc.queue.update(bu, {"status": "rejected", "correct_category": "booking", "correct_hotel": "Hotel Beta"})
    proc.queue.update(ph, {"status": "approved"})
    proc.queue.update(med, {"status": "rejected"})
    proc.process_queue(cfg)

    corrected = proc.queue.get(bu)
    assert corrected["status"] == "pending" and corrected["is_corrected"] is True
    assert corrected["target_email"] == "beta@example.com" and corrected["auto_approved"] is False
    assert json.loads((base_dir / "corrections.json").read_text(encoding="utf-8"))[0]["correct_hotel"] == "Hotel Beta"
    assert ("Rechnung MOCK_SONSTIGES", "EmailProcessor/Sonstiges") in backend.moves
    assert ("Bewerbung MOCK_MEDIUM", "EmailProcessor/Abgelehnt") in backend.moves

    proc.queue.update(bu, {"status": "approved"})
    proc.process_queue(cfg)
    assert any(s[0] == "forward" and s[1] == "beta@example.com" for s in backend.sent)

    records, summary = merge_events(list((base_dir / "training_data").glob("emails_*.jsonl")))
    assert len(records) == 5
    r_bu = next(r for r in records if r["id"] == bu)
    assert (r_bu["label"], r_bu["hotel"], r_bu["label_quality"], r_bu["was_corrected"]) == \
        ("booking", "Hotel Beta", "gold", True)
    assert next(r for r in records if r["id"] == ph)["label"] == "security_threat"
    assert summary["count"] == {"gold": 3, "silver": 2, "rejected": 1, "ignored": 0, "open": 1}
    assert summary["ai"] == {"total": 3, "ok": 2}


def test_shadow_mode_only_marks(base_dir, mock_server):
    cfg = load_config(base_dir)
    cfg["auto_approve"]["mode"] = "shadow"
    backend = FakeBackend()
    backend.add(make_mail("Zimmer B MOCK_BOOKING_ALPHA", "Alpha bitte"))
    proc = Processor(base_dir, backend)
    proc.scan_inbox(cfg)
    proc.process_queue(cfg)
    e = proc.queue.load()[0]
    assert e["status"] == "pending" and e["auto_eligible"] is True and e["auto_mode"] == "shadow"
    assert backend.sent == []


def test_wrong_key_falls_back_to_ollama(base_dir, mock_server, monkeypatch, caplog):
    monkeypatch.setenv("LANGDOCK_API_KEY", "falsch")
    backend = FakeBackend()
    backend.add(make_mail("Angebot C MOCK_BOOKING_ALPHA", "Alpha"))
    proc = Processor(base_dir, backend)
    start = len(mock_server.requests)
    proc.scan_inbox(load_config(base_dir))
    e = proc.queue.load()[0]
    assert e["llm_provider"] == "ollama" and e["category"] == "advertisement" and e["status"] == "pending"
    assert e["auto_approved"] is False
    assert len([r for r in requests_since(mock_server, start) if r["path"] == "/api/chat"]) >= 2
    assert "HTTP 401" in caplog.text


def test_wrong_key_without_fallback_leaves_mail(base_dir, monkeypatch):
    monkeypatch.setenv("LANGDOCK_API_KEY", "falsch")
    cfg = load_config(base_dir)
    cfg["langdock"]["fallback_to_ollama"] = False
    backend = FakeBackend()
    backend.add(make_mail("Angebot D"))
    proc = Processor(base_dir, backend)
    proc.scan_inbox(cfg)
    assert proc.queue.load() == [] and backend.moves == []


def test_ollama_only_mode(base_dir):
    cfg = load_config(base_dir)
    cfg["llm_provider"] = "ollama"
    cfg["training_data"]["enabled"] = False
    backend = FakeBackend()
    backend.add(make_mail("Angebot E"))
    proc = Processor(base_dir, backend)
    proc.scan_inbox(cfg)
    e = proc.queue.load()[0]
    assert (e["llm_provider"], e["confidence"], e["llm_reason"]) == ("ollama", "high", "Ollama sagt Werbung")
    assert not (base_dir / "training_data").exists()


def test_already_known_mails_are_skipped(base_dir):
    cfg = load_config(base_dir)
    backend = FakeBackend()
    backend.add(make_mail("Rechnung MOCK_SONSTIGES"))
    proc = Processor(base_dir, backend)
    proc.scan_inbox(cfg)
    proc.scan_inbox(cfg)  # Mail liegt jetzt im Prüf-Ordner -> kein zweiter Eintrag
    assert len(proc.queue.load()) == 1


def test_missing_mail_on_execute_gives_error(base_dir):
    cfg = load_config(base_dir)
    proc = Processor(base_dir, FakeBackend())
    proc.queue.add({"id": "x1", "entry_id": "gibt-es-nicht", "subject": "Weg", "category": "advertisement",
                    "status": "approved", "target_folder": "EmailProcessor/Werbung"})
    proc.process_queue(cfg)
    assert proc.queue.get("x1")["status"] == "error"
