"""Ablauf: Posteingang scannen, Mails einordnen, Warteschlange abarbeiten."""
from __future__ import annotations

import uuid
from datetime import date, datetime, time, timedelta
from pathlib import Path

import requests

from . import autoapprove, training
from .actions import execute_entry
from .entries import build_queue_entry, corrected_fields, hotels_block
from .llm import langdock, ollama
from .log import log
from .mail.base import MailBackend
from .models import Mail
from .security import check_rules, is_suspicious_spam, merge_llm_phishing, should_skip
from .store import CorrectionsStore, QueueStore, StatsStore

KNOWN_STATUSES = {"pending", "approved", "rejected", "done", "open_task", "task_done", "ignored", "dismissed"}


def llm_provider(cfg: dict) -> str:
    return "langdock" if str(cfg.get("llm_provider") or "").lower() == "langdock" else "ollama"


class Processor:
    def __init__(self, base_dir: Path, backend: MailBackend, http=None) -> None:
        self.base_dir = Path(base_dir)
        self.backend = backend
        self.http = http or requests.Session()
        self.queue = QueueStore(self.base_dir / "queue.json")
        self.stats = StatsStore(self.base_dir / "stats.json")
        self.corrections = CorrectionsStore(self.base_dir / "corrections.json")
        self.training = training.TrainingLog(self.base_dir)

    # ── KI ────────────────────────────────────────────────────────────
    def langdock_analysis(self, mail: Mail, cfg: dict) -> dict:
        """Eine Langdock-Abfrage: Kategorie, Hotel, Sicherheit, Sprache, Phishing. Wirft bei Fehlern."""
        max_chars = int((cfg.get("langdock") or {}).get("body_chars") or 3000)
        prompt = langdock.analysis_prompt(
            hotels_block(cfg), self.corrections.examples_block(), mail.sender_name, mail.sender_email,
            mail.subject, ", ".join(mail.attachments) or "keine", (mail.body or "").strip()[:max_chars])
        result = langdock.normalize_result(
            langdock.invoke_json(prompt, langdock.analysis_schema(), cfg, http=self.http), cfg)
        log.info("  -> Langdock: %s | Hotel: %s | %s | %s",
                 result["category"], result["hotel"], result["confidence"], result["reason"])
        if result["is_phishing"]:
            log.warning("  -> Langdock Phishing-Verdacht (%s): %s",
                        result["phishing_confidence"], result["phishing_reason"])
        return result

    def security_with_ollama(self, mail: Mail, cfg: dict) -> dict:
        sec = check_rules(mail)
        if sec["threat_level"] != "danger" and cfg.get("ollama_model"):
            res = ollama.phishing_check(mail, cfg, http=self.http)
            if res:
                merge_llm_phishing(sec, res["is_phishing"], res["confidence"], res["reason"])
        return sec

    def ollama_classify(self, mail: Mail, cfg: dict) -> dict:
        return ollama.classify(mail, cfg, hotels_block(cfg), self.corrections.examples_block(), http=self.http)

    # ── Posteingang ───────────────────────────────────────────────────
    def scan_inbox(self, cfg: dict) -> None:
        self.stats.set_last_checked()
        known_ids: set[str] = set()
        known_msg_ids: set[str] = set()
        for e in self.queue.load():
            if e.get("status") in KNOWN_STATUSES:
                if e.get("entry_id"):
                    known_ids.add(e["entry_id"])
                if e.get("message_id"):
                    known_msg_ids.add(e["message_id"])

        max_age = int(cfg.get("scan_max_age_days") or 1)
        max_per = int(cfg.get("scan_max_per_run") or 50)
        cutoff = datetime.combine(date.today() - timedelta(days=max_age - 1), time.min)
        mails = [m for m in self.backend.list_new(cutoff, max_per, known_ids)
                 if not (m.message_id and m.message_id in known_msg_ids)]
        if not mails:
            log.info("Keine neuen E-Mails (letzte %s Tage).", max_age)
            return

        log.info("=" * 55)
        log.info("%s neue E-Mail(s) - werden zur Pruefung eingereiht", len(mails))
        log.info("=" * 55)

        use_langdock = llm_provider(cfg) == "langdock"
        fallback = (cfg.get("langdock") or {}).get("fallback_to_ollama") is not False
        processed: set[str] = set()  # Duplikat-Schutz: Absender+Betreff innerhalb dieses Scans

        for i, mail in enumerate(mails, start=1):
            tag = f"[{i}/{len(mails)}]"
            key = f"{(mail.sender_email or '').lower()}|{mail.subject}"
            if key in processed:
                log.info('Duplikat uebersprungen: "%s" | %s', mail.subject, mail.sender_name)
                continue

            # Interne/Kollegen-Emails: rot markieren und als offene Aufgabe speichern
            if not cfg.get("test_mode") and should_skip(mail, cfg):
                self._colleague_task(mail)
                processed.add(key)
                continue

            # Genehmigte Einträge vor dem KI-Aufruf abarbeiten
            self.process_queue(cfg)

            train_input = training.TrainingLog.input_for(mail, cfg) if training.enabled(cfg) else None
            prediction: dict | None = None
            extra: dict = {}

            # Sicherheit: erst Regeln, dann EINE Langdock-Abfrage (liefert auch die Phishing-Einschätzung)
            if use_langdock:
                sec = check_rules(mail)
                if sec["threat_level"] != "danger":
                    log.info('%s Analysiere mit Langdock: "%s" | %s', tag, mail.subject, mail.sender_name)
                    try:
                        prediction = self.langdock_analysis(mail, cfg)
                    except Exception as exc:
                        log.error("  -> Langdock-Fehler: %s", exc)
                    if prediction:
                        merge_llm_phishing(sec, prediction["is_phishing"], prediction["phishing_confidence"],
                                           prediction["phishing_reason"])
                    elif fallback:
                        log.warning("  -> Weiter mit Ollama (Fallback)")
                        sec = self.security_with_ollama(mail, cfg)
                    else:
                        log.warning("  -> Bleibt im Posteingang, neuer Versuch beim naechsten Scan")
                        continue
            else:
                sec = self.security_with_ollama(mail, cfg)

            if sec["is_threat"]:
                self._security_threat(mail, sec, tag, cfg, train_input, prediction, extra)
                processed.add(key)
                continue

            # Verdächtige Spam-Absender: ohne KI direkt als Werbung (bei Langdock-Ergebnis nur als Merkmal)
            spam = is_suspicious_spam(mail)
            extra["spam_heuristic"] = spam
            if spam and not prediction:
                self._spam(mail, tag, cfg, train_input, extra)
                processed.add(key)
                continue

            try:
                if not prediction:
                    log.info('%s Klassifiziere: "%s" | %s', tag, mail.subject, mail.sender_name)
                    cls = self.ollama_classify(mail, cfg)
                    prediction = {"provider": "ollama", "model": str(cfg.get("ollama_model")), **cls, "language": ""}
                elif (cfg.get("training_data") or {}).get("compare_ollama") and train_input:
                    cls = self.ollama_classify(mail, cfg)  # nur zum Vergleich in den Trainingsdaten
                    extra["ollama_prediction"] = {"model": str(cfg.get("ollama_model")), "category": cls["category"],
                                                  "hotel": cls["hotel"], "confidence": cls["confidence"]}

                entry = build_queue_entry(mail, prediction["category"], prediction["hotel"], cfg,
                                          language=prediction.get("language") or "",
                                          detect_language=lambda: ollama.detect_language(mail, cfg, http=self.http))
                entry.update(security_status="clean", message_id=mail.message_id,
                             llm_provider=prediction["provider"], confidence=prediction["confidence"],
                             llm_reason=prediction["reason"])
                autoapprove.apply(entry, prediction, cfg)

                # Erst verschieben, dann neue ID speichern (die ID ändert sich beim Verschieben)
                entry["entry_id"] = self.backend.move(mail.id, "EmailProcessor/Zur Pruefung")
                self.queue.add(entry)
                processed.add(key)
                self.training.prediction(cfg, entry, train_input, prediction, extra)

                if entry["auto_approved"]:
                    log.info("  -> Automatisch freigegeben (%s): %s", entry["auto_note"], entry["proposed_action"])
                elif entry["auto_mode"] == "shadow" and entry["auto_eligible"]:
                    log.info("  -> Testbetrieb: wuerde automatisch ausgefuehrt - wartet auf Genehmigung: %s",
                             entry["proposed_action"])
                else:
                    note = f" (manuell: {entry['auto_note']})" if entry["auto_note"] else ""
                    log.info("  -> Wartet auf Genehmigung: %s%s", entry["proposed_action"], note)
            except Exception as exc:
                log.error("  -> Fehler: %s", exc)

    def _colleague_task(self, mail: Mail) -> None:
        log.info('Kollegen-Email: "%s" | %s', mail.subject, mail.sender_name)
        try:
            self.backend.flag(mail.id)
        except Exception as exc:
            log.warning("  -> Markierung fehlgeschlagen: %s", exc)
        body = (mail.body or "").strip()
        self.queue.add({
            "id": str(uuid.uuid4()),
            "entry_id": mail.id,
            "message_id": mail.message_id,
            "subject": (mail.subject or "")[:80],
            "sender": mail.sender_email,
            "sender_name": mail.sender_name,
            "received": mail.received.strftime("%d.%m.%Y %H:%M"),
            "created": datetime.now().strftime("%d.%m. %H:%M"),
            "category": "task",
            "status": "open_task",
            "proposed_action": "Manuelle Bearbeitung",
            "body_preview": body[:2000] + "…" if len(body) > 2000 else body,
        })
        log.info("  -> Rot markiert + als Aufgabe gespeichert")

    def _security_threat(self, mail, sec, tag, cfg, train_input, prediction, extra) -> None:
        reasons = " | ".join(sec["reasons"])
        log.warning('%s SICHERHEITSWARNUNG (%s): "%s" | %s', tag, sec["threat_level"].upper(), mail.subject,
                    mail.sender_name)
        log.warning("  -> Gruende: %s", reasons)
        try:
            entry = build_queue_entry(mail, "security_threat", "", cfg)
            entry.update(proposed_action=f"SICHERHEITSWARNUNG ({sec['threat_level']}): {reasons}",
                         target_folder="EmailProcessor/Sicherheitswarnung", security_warning=reasons,
                         threat_level=sec["threat_level"], message_id=mail.message_id)
            entry["entry_id"] = self.backend.move(mail.id, "EmailProcessor/Sicherheitswarnung")
            self.queue.add(entry)
            log.info("  -> Verschoben nach: EmailProcessor/Sicherheitswarnung")
            extra.update(security_reasons=list(sec["reasons"]), threat_level=sec["threat_level"])
            self.training.prediction(cfg, entry, train_input, prediction, extra)
        except Exception as exc:
            log.error("  -> Fehler beim Sicherheits-Handling: %s", exc)

    def _spam(self, mail, tag, cfg, train_input, extra) -> None:
        log.info('%s Spam-Absender erkannt: "%s" | %s', tag, mail.subject, mail.sender_name)
        try:
            entry = build_queue_entry(mail, "advertisement", "", cfg)
            entry["proposed_action"] = "Archivieren (verdaechtiger Spam-Absender — kein echter Name gesetzt)"
            entry["message_id"] = mail.message_id
            autoapprove.apply(entry, None, cfg)
            entry["entry_id"] = self.backend.move(mail.id, "EmailProcessor/Zur Pruefung")
            self.queue.add(entry)
            log.info("  -> Als Werbung/Spam eingereiht (KI uebersprungen)")
            self.training.prediction(cfg, entry, train_input, None, extra)
        except Exception as exc:
            log.error("  -> Fehler beim Spam-Handling: %s", exc)

    # ── Warteschlange ─────────────────────────────────────────────────
    def process_queue(self, cfg: dict) -> None:
        """Genehmigte, ignorierte und abgelehnte Einträge abarbeiten.

        Die Aktionen laufen ohne Sperre (sie können dauern); danach wird nur der jeweilige Eintrag
        aktualisiert - Änderungen, die das Dashboard inzwischen gemacht hat, bleiben erhalten.
        """
        todo = [e for e in self.queue.load() if e.get("status") in ("approved", "ignored", "rejected")]
        for entry in todo:
            status = entry.get("status")
            if status == "approved":
                how = "Automatisch freigegeben" if entry.get("auto_approved") else "Genehmigt"
                log.info('%s - fuehre aus: "%s"', how, entry.get("subject"))
                result = execute_entry(entry, self.backend, self.http)
                self.queue.update(entry["id"], {"status": result})
                entry["status"] = result
                self.stats.record(entry, result)
                self.training.decision(cfg, entry, "auto_approved" if entry.get("auto_approved") else "approved",
                                       entry.get("category") or "", entry.get("hotel") or "", result)
            elif status == "ignored":
                log.info('Ignoriert: "%s"', entry.get("subject"))
                self.training.decision(cfg, entry, "ignored")
                try:
                    self.backend.move(entry.get("entry_id") or "", "EmailProcessor/Ignoriert")
                    log.info("  -> Verschoben nach EmailProcessor/Ignoriert")
                except Exception:
                    log.warning("  -> E-Mail nicht gefunden oder bereits verschoben")
                self.queue.update(entry["id"], {"status": "done"})
            elif status == "rejected":
                self._rejected(entry, cfg)

    def _rejected(self, entry: dict, cfg: dict) -> None:
        reason = f" | {entry['rejection_reason']}" if entry.get("rejection_reason") else ""
        log.info('Abgelehnt%s: "%s"', reason, entry.get("subject"))
        correct = entry.get("correct_category")
        if correct and correct not in ("unknown", ""):
            # Korrektur: merken, Eintrag anpassen und erneut zur Genehmigung stellen
            hotel_info = f" | {entry['correct_hotel']}" if entry.get("correct_hotel") else ""
            log.info("  -> Korrektur: %s%s", correct, hotel_info)
            self.corrections.save_correction(entry)
            log.info("  -> Korrektur gespeichert: %s -> %s", entry.get("category"), correct)
            self.training.decision(cfg, entry, "corrected", correct, entry.get("correct_hotel") or "")
            fields = {**entry, **corrected_fields(entry, cfg)}
            update = {k: fields.get(k) for k in ("category", "hotel", "target_folder", "target_email",
                                                 "proposed_action", "action_type", "reply_to", "reply_draft")}
            update.update(is_corrected=True, auto_eligible=False, auto_approved=False,
                          auto_note="korrigierter Eintrag", correct_category=None, correct_hotel=None,
                          rejection_reason=None, status="pending")
            self.queue.update(entry["id"], update)
            log.info("  -> Zurueck in Warteschlange fuer erneute Genehmigung")
        else:
            self.training.decision(cfg, entry, "rejected")
            try:
                self.backend.move(entry.get("entry_id") or "", "EmailProcessor/Abgelehnt")
                log.info("  -> Archiviert in EmailProcessor/Abgelehnt")
            except Exception:
                log.warning("  -> E-Mail nicht gefunden")
            self.queue.update(entry["id"], {"status": "done"})
