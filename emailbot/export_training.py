"""Führt die gesammelten Trainingsdaten zusammen: ein Datensatz pro Mail.
Zeigt außerdem, wie oft die KI richtig lag und wie gut die Auto-Freigabe gewesen wäre.

    python -m emailbot.export_training
    python -m emailbot.export_training --dir D:\\daten\\training_data --out D:\\export.jsonl
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import date
from pathlib import Path

from .config import DEFAULT_BASE_DIR, DEFAULTS, deep_merge, read_json
from .training import data_dir


def merge_events(files: list[Path]) -> tuple[list[dict], dict]:
    preds: dict[str, dict] = {}
    decs: dict[str, list[dict]] = {}
    order: list[str] = []
    bad = 0
    for f in sorted(files):
        with open(f, "r", encoding="utf-8-sig") as fh:
            for line in fh:
                if not line.strip():
                    continue
                try:
                    ev = json.loads(line)
                except json.JSONDecodeError:
                    bad += 1
                    continue
                eid = ev.get("id")
                if not eid:
                    bad += 1
                    continue
                if ev.get("event") == "prediction":
                    if eid not in preds:
                        order.append(eid)
                    preds[eid] = ev
                elif ev.get("event") == "decision":
                    decs.setdefault(eid, []).append(ev)

    count = {"gold": 0, "silver": 0, "rejected": 0, "ignored": 0, "open": 0}
    labels: dict[str, int] = {}
    ev_ai = {"total": 0, "ok": 0}
    ev_hotel = {"total": 0, "ok": 0}
    ev_shadow = {"total": 0, "ok": 0}
    ev_ollama = {"total": 0, "ok": 0}
    records: list[dict] = []

    for eid in order:
        p = preds[eid]
        dl = decs.get(eid, [])
        if not dl:
            count["open"] += 1
            continue
        last = dl[-1]
        corrected = any(d.get("decision") == "corrected" for d in dl)
        label = hotel = None
        quality = {"approved": "gold", "corrected": "gold", "auto_approved": "silver",
                   "rejected": "rejected", "ignored": "ignored"}.get(last.get("decision"), "open")
        if quality in ("gold", "silver"):
            label, hotel = last.get("final_category"), last.get("final_hotel")
        count[quality] += 1

        auto = p.get("auto") or {}
        if auto.get("eligible") is True and quality in ("gold", "rejected", "ignored"):
            ev_shadow["total"] += 1
            if quality == "gold" and not corrected:
                ev_shadow["ok"] += 1

        if not label:
            continue
        labels[label] = labels.get(label, 0) + 1
        if quality == "gold":
            if p.get("queued_category"):
                ev_ai["total"] += 1
                if p["queued_category"] == label and not corrected:
                    ev_ai["ok"] += 1
            if label == "booking" and hotel:
                ev_hotel["total"] += 1
                if (p.get("queued_hotel") or "") == hotel:
                    ev_hotel["ok"] += 1
            op = (p.get("extra") or {}).get("ollama_prediction")
            if op:
                ev_ollama["total"] += 1
                if op.get("category") == label:
                    ev_ollama["ok"] += 1

        records.append({
            "id": eid,
            "message_id": (p.get("input") or {}).get("message_id"),
            "received": (p.get("input") or {}).get("received"),
            "input": p.get("input"),
            "label": label,
            "hotel": hotel,
            "label_quality": quality,
            "was_corrected": corrected,
            "prediction": p.get("prediction"),
            "ollama_prediction": (p.get("extra") or {}).get("ollama_prediction"),
            "queued_category": p.get("queued_category"),
            "auto": p.get("auto"),
            "prompt_version": p.get("prompt_version"),
            "decided_at": last.get("time"),
        })

    summary = {"mails": len(order), "count": count, "labels": labels, "bad": bad,
               "ai": ev_ai, "hotel": ev_hotel, "shadow": ev_shadow, "ollama": ev_ollama}
    return records, summary


def _rate(e: dict) -> str:
    if not e["total"]:
        return "noch keine Daten"
    return f"{e['ok']} von {e['total']} ({e['ok'] / e['total']:.0%})"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="python -m emailbot.export_training")
    parser.add_argument("--dir", help="Ordner mit emails_*.jsonl (Standard aus config.json)")
    parser.add_argument("--out", help="Ausgabedatei (Standard: <dir>/export_DATUM.jsonl)")
    args = parser.parse_args(argv)
    try:
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass

    if args.dir:
        d = Path(args.dir)
    else:
        cfg = deep_merge(DEFAULTS, read_json(DEFAULT_BASE_DIR / "config.json", default={}) or {})
        d = data_dir(DEFAULT_BASE_DIR, cfg)
    if not d.exists():
        print(f"Ordner nicht gefunden: {d}")
        return 1
    files = list(d.glob("emails_*.jsonl"))
    if not files:
        print(f"Keine Dateien emails_*.jsonl in {d}")
        return 0
    out = Path(args.out) if args.out else d / f"export_{date.today():%Y-%m-%d}.jsonl"

    records, s = merge_events(files)
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        for r in records:
            f.write(json.dumps(r, ensure_ascii=False, separators=(",", ":")) + "\n")

    c = s["count"]
    print("\nTrainingsdaten zusammengefuehrt")
    print(f"  Datei:            {out}")
    print(f"  Mails gesamt:     {s['mails']}")
    print(f"  gold (von dir bestaetigt/korrigiert): {c['gold']}")
    print(f"  silver (automatisch ausgefuehrt):     {c['silver']}")
    print(f"  abgelehnt ohne Korrektur:             {c['rejected']}")
    print(f"  ignoriert:                            {c['ignored']}")
    print(f"  noch offen:                           {c['open']}")
    if s["bad"]:
        print(f"  unlesbare Zeilen:                     {s['bad']}")
    print("\nDatensaetze je Kategorie:")
    for k in sorted(s["labels"]):
        print(f"  {k:<16} {s['labels'][k]}")
    print("\nWie oft lag die KI richtig? (nur von dir gepruefte Mails)")
    print(f"  Kategorie:              {_rate(s['ai'])}")
    print(f"  Hotel bei Buchungen:    {_rate(s['hotel'])}")
    print(f"  Auto-Freigabe (Test):   {_rate(s['shadow'])}")
    if s["ollama"]["total"]:
        print(f"  Ollama zum Vergleich:   {_rate(s['ollama'])}")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
