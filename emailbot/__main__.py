"""Startpunkt: Bot und Dashboard in einem Prozess.

    python -m emailbot              Bot + Dashboard (http://localhost:5000)
    python -m emailbot --paused     Dashboard starten, Bot erst per Knopf im Dashboard
    python -m emailbot --no-dashboard
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import DEFAULT_BASE_DIR, load_config
from .log import log, setup_logging
from .mail import make_backend, open_in_mail_client
from .runner import BotRunner


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="python -m emailbot", description="E-Mail-Bot mit Dashboard")
    parser.add_argument("--dir", help="Programmordner mit config.json (Standard: Ordner über dem Paket)")
    parser.add_argument("--port", type=int, help="Port des Dashboards (Standard aus config.json, sonst 5000)")
    parser.add_argument("--no-dashboard", action="store_true", help="nur den Bot starten")
    parser.add_argument("--paused", action="store_true", help="Bot angehalten starten")
    args = parser.parse_args(argv)

    base_dir = Path(args.dir).resolve() if args.dir else DEFAULT_BASE_DIR
    setup_logging(base_dir)
    try:
        cfg = load_config(base_dir)
    except Exception as exc:
        log.error("%s", exc)
        return 1

    runner = BotRunner(base_dir, make_backend, start_paused=args.paused)

    if args.no_dashboard:
        try:
            runner.run()
        except KeyboardInterrupt:
            pass
        return 0

    from werkzeug.serving import make_server
    from .web import create_app

    port = args.port or int(cfg.get("dashboard_port") or 5000)
    app = create_app(base_dir, runner, open_mail=open_in_mail_client)
    try:
        # Nur localhost: das Dashboard ist von anderen Rechnern aus nicht erreichbar
        server = make_server("127.0.0.1", port, app, threaded=True)
    except OSError:
        log.error("Port %s ist belegt - laeuft der Bot (oder dashboard_server.ps1) schon?", port)
        return 1

    runner.start()
    log.info("Dashboard laeuft auf http://localhost:%s  (Strg+C zum Beenden)", port)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        runner.stop()
        server.shutdown()
    return 0


if __name__ == "__main__":
    sys.exit(main())
