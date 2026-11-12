"""Logging im selben Format wie die PowerShell-Version (email_processor.log)."""
from __future__ import annotations

import logging
import sys
from pathlib import Path

log = logging.getLogger("emailbot")


def setup_logging(base_dir: Path, console: bool = True) -> None:
    logging.addLevelName(logging.WARNING, "WARN")
    log.setLevel(logging.INFO)
    log.propagate = False
    for handler in list(log.handlers):
        log.removeHandler(handler)
        handler.close()

    fmt = logging.Formatter("%(asctime)s %(levelname)-8s %(message)s", "%Y-%m-%d %H:%M:%S")
    file_handler = logging.FileHandler(Path(base_dir) / "email_processor.log", encoding="utf-8")
    file_handler.setFormatter(fmt)
    log.addHandler(file_handler)

    if console:
        try:
            sys.stdout.reconfigure(errors="replace")  # Umlaute in alten Konsolen nicht abstürzen lassen
        except (AttributeError, ValueError):
            pass
        stream = logging.StreamHandler(sys.stdout)
        stream.setFormatter(fmt)
        log.addHandler(stream)
