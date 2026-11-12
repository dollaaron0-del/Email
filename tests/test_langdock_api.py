"""Langdock-Aufrufe gegen den Testserver."""
import json

import pytest

from emailbot.llm import langdock

from .conftest import ALPHA_REASON


def last_request(server):
    return server.requests[-1]


def test_completion_api_request_and_umlauts(base_cfg, mock_server):
    res = langdock.invoke_json("MOCK_BOOKING_ALPHA Testprompt", langdock.analysis_schema(), base_cfg)
    assert res["category"] == "booking" and res["hotel_name"] == "Hotel Alpha"
    assert res["reason"] == ALPHA_REASON  # Antwort ohne charset -> trotzdem korrekt als UTF-8 gelesen
    req = last_request(mock_server)
    sent = json.loads(req["body"])
    assert req["path"] == "/openai/eu/v1/chat/completions"
    assert req["auth"] == "Bearer testkey"
    assert sent["model"] == "test-model"
    assert sent["response_format"] == {"type": "json_object"}
    assert sent["messages"][0]["content"] == "MOCK_BOOKING_ALPHA Testprompt"
    assert "temperature" not in sent


def test_optional_parameters_are_sent(base_cfg, mock_server):
    base_cfg["langdock"].update(temperature=0.2, reasoning_effort="low")
    langdock.invoke_json("x", None, base_cfg)
    sent = json.loads(last_request(mock_server)["body"])
    assert sent["temperature"] == 0.2 and sent["reasoning_effort"] == "low"


def test_agent_api_with_structured_output(base_cfg, mock_server):
    base_cfg["langdock"].update(model="", agent_id="agent-123")
    res = langdock.invoke_json("MOCK_NEWSLETTER", langdock.analysis_schema(), base_cfg)
    sent = json.loads(last_request(mock_server)["body"])
    assert res["category"] == "newsletter"
    assert last_request(mock_server)["path"] == "/agent/v1/chat/completions"
    assert sent["agentId"] == "agent-123" and sent["stream"] is False
    assert sent["messages"][0]["parts"][0]["text"] == "MOCK_NEWSLETTER"
    assert sent["output"]["type"] == "object"


def test_agent_api_without_schema_reads_message_content(base_cfg):
    base_cfg["langdock"].update(model="", agent_id="agent-123")
    assert langdock.invoke_json("MOCK_SONSTIGES", None, base_cfg)["category"] == "sonstiges"


def test_401_message_without_key(base_cfg, monkeypatch):
    monkeypatch.setenv("LANGDOCK_API_KEY", "falsch-geheim")
    with pytest.raises(langdock.LangdockError) as err:
        langdock.invoke_json("x", None, base_cfg)
    msg = str(err.value)
    assert "HTTP 401" in msg and "invalid api key" in msg and "falsch-geheim" not in msg


def test_missing_key(base_cfg, monkeypatch):
    monkeypatch.delenv("LANGDOCK_API_KEY")
    with pytest.raises(langdock.LangdockError, match="Kein Langdock-API-Key"):
        langdock.invoke_json("x", None, base_cfg)


def test_missing_model(base_cfg):
    base_cfg["langdock"]["model"] = ""
    with pytest.raises(langdock.LangdockError, match="langdock.model fehlt"):
        langdock.invoke_json("x", None, base_cfg)


def test_connection_error_is_langdock_error(base_cfg):
    base_cfg["langdock"]["base_url"] = "http://127.0.0.1:9"  # nichts erreichbar
    base_cfg["langdock"]["timeout_sec"] = 3
    with pytest.raises(langdock.LangdockError, match="Verbindung zu Langdock fehlgeschlagen"):
        langdock.invoke_json("x", None, base_cfg)


def test_list_models(base_cfg):
    assert langdock.list_models(base_cfg) == ["other-model", "test-model"]


def test_key_from_registry_when_env_missing(base_cfg, monkeypatch):
    monkeypatch.delenv("LANGDOCK_API_KEY")
    monkeypatch.setattr(langdock, "_read_registry_env", lambda name: " aus-registry ")
    assert langdock.get_api_key(base_cfg) == "aus-registry"
