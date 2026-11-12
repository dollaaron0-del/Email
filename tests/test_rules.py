"""Reine Regeln: Sicherheit, Abmelde-Header, JSON-Parser, Hotel-Suche, Auto-Freigabe."""
import pytest

from emailbot import autoapprove
from emailbot.config import DEFAULTS, deep_merge
from emailbot.entries import find_hotel
from emailbot.llm import parse_llm_json
from emailbot.llm.langdock import analysis_schema, normalize_result
from emailbot.security import check_rules, is_suspicious_spam, merge_llm_phishing, should_skip, unsubscribe_info

from .conftest import make_mail


def test_dangerous_attachment_is_danger():
    sec = check_rules(make_mail("x", attachments=["rechnung.exe"]))
    assert sec["threat_level"] == "danger" and "rechnung.exe" in sec["reasons"][0]


def test_macro_attachment_and_ip_link_warn():
    assert check_rules(make_mail("x", attachments=["a.xlsm"]))["threat_level"] == "warn"
    assert check_rules(make_mail("x", body="siehe http://10.1.2.3/login"))["threat_level"] == "warn"


def test_phishing_url_is_danger():
    assert check_rules(make_mail("x", body="https://paypal-secure.example.net/verify"))["threat_level"] == "danger"


def test_clean_mail():
    assert check_rules(make_mail("x", body="Hallo https://www.althoff.de"))["is_threat"] is False


@pytest.mark.parametrize("conf,level", [("high", "danger"), ("medium", "danger"), ("low", "warn")])
def test_merge_llm_phishing(conf, level):
    sec = merge_llm_phishing({"threat_level": "none", "reasons": [], "is_threat": False}, True, conf, "Bank")
    assert sec["threat_level"] == level and sec["is_threat"]


def test_merge_llm_no_phishing_keeps_clean():
    sec = merge_llm_phishing({"threat_level": "none", "reasons": [], "is_threat": False}, False, "high", "")
    assert not sec["is_threat"]


def test_unsubscribe_folded_header_and_one_click():
    headers = ("Subject: News\r\nList-Unsubscribe: <mailto:off@shop.de?subject=stop>,\r\n"
               "  <https://shop.de/unsub?id=1>\r\nList-Unsubscribe-Post: List-Unsubscribe=One-Click\r\n")
    info = unsubscribe_info(headers)
    assert info["mailto"] == "off@shop.de?subject=stop"
    assert info["http"] == "https://shop.de/unsub?id=1"
    assert info["one_click"] is True


def test_unsubscribe_missing():
    assert unsubscribe_info("Subject: x\r\n") == {"raw": "", "mailto": "", "http": "", "one_click": False}


def test_spam_heuristic():
    assert is_suspicious_spam(make_mail("x", sender="xk29fjs83kd9a@gmail.com", name="Anna"))
    assert is_suspicious_spam(make_mail("x", sender="anna@gmail.com", name="anna"))
    assert not is_suspicious_spam(make_mail("x", sender="anna.schmidt@gmail.com", name="Anna Schmidt"))
    assert not is_suspicious_spam(make_mail("x", sender="info@firma.de", name="info@firma.de"))


def test_should_skip():
    cfg = {"skip_domains": ["meine-firma.de"], "skip_senders": ["chef@extern.de"]}
    assert should_skip(make_mail("x", internal=True), cfg)
    assert should_skip(make_mail("x", sender="kollege@MEINE-FIRMA.de"), cfg)
    assert should_skip(make_mail("x", sender="chef@extern.de"), cfg)
    assert not should_skip(make_mail("x", sender="gast@gmail.com"), cfg)


def test_parse_llm_json():
    assert parse_llm_json('Ergebnis:\n```json\n{"category":"booking"}\n```')["category"] == "booking"
    with pytest.raises(ValueError):
        parse_llm_json("kein json")


def test_normalize_result():
    r = normalize_result({"category": "Werbung", "hotel_name": "null", "confidence": "HIGH", "language": "fr",
                          "is_phishing": "true", "phishing_confidence": "?"}, {"langdock": {"model": "m"}})
    assert (r["category"], r["hotel"], r["confidence"], r["language"], r["is_phishing"], r["phishing_confidence"]) == \
        ("unknown", "", "high", "other", True, "low")


def test_schema_allows_null_hotel():
    assert analysis_schema()["properties"]["hotel_name"]["type"] == ["string", "null"]


def test_find_hotel():
    cfg = {"hotels": [{"name": "Hotel Alpha", "email": "a@x", "keywords": ["alpha"]},
                      {"name": "Beta Resort", "email": "b@x", "keywords": ["", "seeblick"]}]}
    assert find_hotel("hotel alpha", cfg)["email"] == "a@x"
    assert find_hotel("Anfrage Seeblick", cfg)["email"] == "b@x"
    assert find_hotel("", cfg) is None and find_hotel("null", cfg) is None
    assert find_hotel("Gamma", cfg) is None  # leeres Keyword trifft nicht alles


def test_deep_merge_keeps_defaults_for_null_sections():
    cfg = deep_merge(DEFAULTS, {"langdock": None, "auto_approve": {"mode": "on"}})
    assert cfg["langdock"]["region"] == "eu" and cfg["auto_approve"]["min_confidence"] == "high"


# ── Auto-Freigabe ─────────────────────────────────────────────────────
def E(**kw):
    e = {"category": "booking", "action_type": "forward", "target_email": "a@b.c", "security_status": "clean",
         "unsubscribe_one_click": False, "is_corrected": False}
    e.update(kw)
    return e


def P(conf="high", provider="langdock"):
    return {"provider": provider, "confidence": conf}


def C(mode="shadow", **aa):
    return {"auto_approve": {"mode": mode, "min_confidence": "high",
                             "categories": ["advertisement", "newsletter", "application", "booking"], **aa}}


@pytest.mark.parametrize("entry,pred,cfg,expected", [
    (E(), P(), C("off"), False),
    (E(), P(), C(), True),
    (E(action_type="reply", target_email=""), P(), C(), False),
    (E(category="newsletter"), P(), C(), False),
    (E(category="newsletter", unsubscribe_one_click=True), P(), C(), True),
    (E(category="sonstiges"), P(), C(), False),
    (E(category="security_threat"), P(), C(), False),
    (E(), P("medium"), C(), False),
    (E(), P("medium"), C(min_confidence="medium"), True),
    (E(), P("low"), C(min_confidence="medium"), False),
    (E(), P(provider="ollama"), C(), False),
    (E(), P(provider="ollama"), C(allow_ollama=True), True),
    (E(security_status=""), P(), C(), False),
    (E(category="application"), P(), C(categories=["booking"]), False),
    (E(is_corrected=True), P(), C(), False),
    (E(), None, C(), False),
])
def test_auto_decide(entry, pred, cfg, expected):
    assert autoapprove.decide(entry, pred, cfg)[0] is expected


def test_auto_test_mode_blocks():
    assert autoapprove.decide(E(), P(), {**C(), "test_mode": True}) == (False, "Test-Modus aktiv")


def test_auto_apply_modes():
    on = E(status="pending")
    autoapprove.apply(on, P(), C("on"))
    assert on["status"] == "approved" and on["auto_approved"] is True
    shadow = E(status="pending")
    autoapprove.apply(shadow, P(), C("shadow"))
    assert shadow["status"] == "pending" and shadow["auto_eligible"] is True and shadow["auto_approved"] is False
