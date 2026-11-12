"""Prüft die Langdock-Verbindung: API-Key, verfügbare Modelle und eine Test-Analyse.

Es wird KEINE echte E-Mail gelesen oder verschickt - nur ein erfundener Beispieltext an Langdock gesendet.

    python -m emailbot.test_langdock
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from .config import DEFAULT_BASE_DIR, deep_merge, DEFAULTS, read_json
from .entries import hotels_block
from .llm import langdock


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="python -m emailbot.test_langdock")
    parser.add_argument("--config", help="Pfad zur config.json (Standard: config.json, sonst config.example.json)")
    args = parser.parse_args(argv)
    try:
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass

    path = Path(args.config) if args.config else DEFAULT_BASE_DIR / "config.json"
    if not path.exists() and not args.config:
        path = DEFAULT_BASE_DIR / "config.example.json"
    cfg = deep_merge(DEFAULTS, read_json(path, default={}) or {})
    print(f"\nConfig: {path}")

    name = langdock.key_name(cfg)
    if not langdock.get_api_key(cfg):
        print(f"FEHLER: Kein API-Key in der Umgebungsvariable {name} gefunden.")
        print("Setzen mit (PowerShell):")
        print(f"  [Environment]::SetEnvironmentVariable('{name}', 'DEIN-KEY', 'User')")
        return 1
    print(f"1) API-Key gefunden ({name})")

    ld = cfg["langdock"]
    if ld.get("agent_id"):
        print(f"2) Agent-Modus: Agent {ld['agent_id']} (Modellliste wird uebersprungen)")
    else:
        try:
            models = langdock.list_models(cfg)
            print("2) Verfuegbare Modelle:")
            for m in models:
                print(f"     {m}")
            if ld.get("model") and ld["model"] not in models:
                print(f"   WARNUNG: Das eingestellte Modell '{ld['model']}' ist nicht in der Liste.")
        except Exception as exc:
            print(f"2) Modellliste nicht abrufbar: {exc}")

    hotel = (cfg.get("hotels") or [{}])[0].get("name") or "Beispiel Hotel"
    prompt = langdock.analysis_prompt(
        hotels_block(cfg), "", "Max Mustermann", "max.mustermann@example.com", "Zimmeranfrage März", "keine",
        f"Guten Tag,\n\nwir möchten vom 12. bis 14. März ein Doppelzimmer im {hotel} buchen. "
        "Ist noch etwas frei?\n\nViele Grüße\nMax Mustermann")
    print(f"3) Test-Analyse (Buchungsanfrage fuer '{hotel}')...")
    started = time.monotonic()
    try:
        res = langdock.normalize_result(langdock.invoke_json(prompt, langdock.analysis_schema(), cfg), cfg)
    except Exception as exc:
        print(f"   FEHLER: {exc}")
        return 1
    print(f"   Antwort nach {time.monotonic() - started:.1f} Sek.:")
    for k, v in res.items():
        print(f"     {k:<20} {v}")
    print("   OK - als Buchung erkannt." if res["category"] == "booking" else "   Hinweis: erwartet war 'booking'.")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
