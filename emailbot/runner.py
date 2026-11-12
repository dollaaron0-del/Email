"""Bot-Schleife im Hintergrund-Thread: alle 10 Sek. Warteschlange, alle N Minuten Posteingang."""
from __future__ import annotations

import threading
import time
from pathlib import Path
from typing import Callable

from . import autoapprove, training
from .config import load_config
from .llm import langdock
from .log import log
from .mail.base import MailBackend
from .processor import Processor, llm_provider

QUEUE_INTERVAL_SEC = 10


def log_startup(cfg: dict) -> None:
    if llm_provider(cfg) == "langdock":
        key = "API-Key gefunden" if langdock.get_api_key(cfg) else f"KEIN API-Key ({langdock.key_name(cfg)})"
        fallback = "aus" if (cfg.get("langdock") or {}).get("fallback_to_ollama") is False else cfg.get("ollama_model")
        log.info("KI: Langdock (%s) | %s | Ollama-Fallback: %s", langdock.model_label(cfg), key, fallback)
    else:
        log.info("KI: Ollama (%s)", cfg.get("ollama_model"))
    log.info("Mailzugriff: %s | Auto-Freigabe: %s | Trainingsdaten: %s", cfg.get("mail_backend"),
             autoapprove.mode(cfg), "an" if training.enabled(cfg) else "aus")
    log.info("Scanintervall: %s Min. | Queue-Check: alle %s Sek.", cfg.get("check_interval_minutes"),
             QUEUE_INTERVAL_SEC)


class BotRunner:
    def __init__(self, base_dir: Path, backend_factory: Callable[[dict], MailBackend], start_paused: bool = False):
        self.base_dir = Path(base_dir)
        self.backend_factory = backend_factory
        self._paused = threading.Event()
        self._stop = threading.Event()
        self._wake = threading.Event()
        self._thread: threading.Thread | None = None
        if start_paused:
            self._paused.set()

    # ── Steuerung (auch aus dem Dashboard) ────────────────────────────
    @property
    def running(self) -> bool:
        return self._thread is not None and self._thread.is_alive() and not self._paused.is_set()

    def start(self) -> None:
        self._thread = threading.Thread(target=self.run, name="emailbot", daemon=True)
        self._thread.start()

    def pause(self) -> None:
        if not self._paused.is_set():
            log.info("Bot angehalten (Dashboard)")
        self._paused.set()

    def resume(self) -> None:
        if self._paused.is_set():
            log.info("Bot fortgesetzt (Dashboard)")
        self._paused.clear()
        self._wake.set()

    def stop(self) -> None:
        self._stop.set()
        self._wake.set()

    def _sleep(self, seconds: float) -> None:
        self._wake.wait(seconds)
        self._wake.clear()

    # ── Schleife ──────────────────────────────────────────────────────
    def run(self) -> None:
        cfg = load_config(self.base_dir)
        log.info("Bot gestartet - Genehmigungen laufen ueber das Dashboard")
        log_startup(cfg)

        backend: MailBackend | None = None
        processor: Processor | None = None
        last_scan = 0.0
        while not self._stop.is_set():
            if self._paused.is_set():
                self._sleep(1)
                continue
            if backend is None:  # erst verbinden, wenn der Bot wirklich laufen soll
                try:
                    backend = self.backend_factory(cfg)
                    backend.connect()
                    processor = Processor(self.base_dir, backend)
                    log.info("Mailzugriff verbunden (%s)", backend.name)
                except NotImplementedError as exc:
                    log.error("%s", exc)
                    self._paused.set()
                    backend = None
                    continue
                except Exception as exc:
                    log.error("Mailzugriff nicht moeglich (%s) - neuer Versuch in 30 Sek.", exc)
                    backend = None
                    self._sleep(30)
                    continue
            try:
                cfg = load_config(self.base_dir)
                backend.ensure_connected()
                processor.process_queue(cfg)
                interval = max(1, int(cfg.get("check_interval_minutes") or 5)) * 60
                if time.monotonic() - last_scan >= interval:
                    processor.scan_inbox(cfg)
                    last_scan = time.monotonic()
            except Exception as exc:
                log.error("Unerwarteter Fehler im Hauptloop: %s", exc)
            self._sleep(QUEUE_INTERVAL_SEC)
        log.info("Bot beendet")
