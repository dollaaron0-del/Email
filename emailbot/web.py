"""Dashboard-Server - dieselben Adressen wie dashboard_server.ps1, die Oberfläche in www/ bleibt gleich."""
from __future__ import annotations

import copy
import json
import time
from datetime import datetime
from pathlib import Path
from typing import Callable

import requests
from flask import Flask, Response, jsonify, request, send_from_directory

from . import autoapprove, training
from .config import DEFAULTS, load_config, read_json, write_json
from .llm import langdock, ollama
from .log import log
from .processor import llm_provider
from .store import QueueStore


def _sort_key(entry: dict) -> datetime:
    try:
        return datetime.strptime(entry.get("received") or "", "%d.%m.%Y %H:%M")
    except ValueError:
        return datetime.min


def create_app(base_dir: Path, runner=None, open_mail: Callable[[dict, str], None] | None = None,
               http=None) -> Flask:
    base_dir = Path(base_dir)
    www = base_dir / "www"
    queue = QueueStore(base_dir / "queue.json")
    app = Flask(__name__, static_folder=None)
    ollama_cache = {"time": 0.0, "ok": False, "url": ""}

    def cfg() -> dict:
        try:
            return load_config(base_dir)
        except Exception:
            return copy.deepcopy(DEFAULTS)  # Dashboard soll auch ohne/mit kaputter config.json laden

    def body() -> dict:
        return request.get_json(silent=True, force=True) or {}

    def set_status(new_status: str, extra: Callable[[dict, dict], None] | None = None):
        data = body()
        with queue.edit() as entries:
            for e in entries:
                if e.get("id") == data.get("id"):
                    e["status"] = new_status
                    if extra:
                        extra(e, data)
                    break
        return jsonify(success=True)

    def ollama_ok(c: dict) -> bool:
        url = ollama.tags_url(c)
        if time.monotonic() - ollama_cache["time"] < 30 and ollama_cache["url"] == url:
            return ollama_cache["ok"]
        try:
            ok = (http or requests).get(url, timeout=2).status_code < 400
        except requests.RequestException:
            ok = False
        ollama_cache.update(time=time.monotonic(), ok=ok, url=url)
        return ok

    # ── Status & Logs ─────────────────────────────────────────────────
    @app.get("/api/stats")
    def api_stats():
        data = read_json(base_dir / "stats.json", default=None) or {}
        c = cfg()
        data.update(
            bot_running=bool(runner and runner.running),
            test_mode=bool(c.get("test_mode")),
            ollama_ok=ollama_ok(c),
            llm_provider=llm_provider(c),
            auto_mode=autoapprove.mode(c),
            langdock_key_set=bool(langdock.get_api_key(c)),  # nur OB es einen Key gibt - nie den Key selbst
            langdock_key_name=langdock.key_name(c),
            training_enabled=training.enabled(c),
            mail_backend=c.get("mail_backend"),
        )
        return jsonify(data)

    @app.get("/api/logs")
    def api_logs():
        path = base_dir / "email_processor.log"
        if not path.exists():
            return jsonify([])
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            lines = f.read().splitlines()
        return jsonify(lines[-40:])

    # ── Warteschlange ─────────────────────────────────────────────────
    @app.get("/api/queue")
    def api_queue():
        items = [e for e in queue.load() if e.get("status") in ("pending", "approved", "rejected")]
        return jsonify(sorted(items, key=_sort_key, reverse=True))

    @app.post("/api/queue/approve")
    def api_approve():
        def extra(e, data):
            if data.get("reply_text"):
                e["reply_text"] = data["reply_text"]
        return set_status("approved", extra)

    @app.post("/api/queue/reject")
    def api_reject():
        def extra(e, data):
            for k in ("rejection_reason", "correct_category", "correct_hotel"):
                if data.get(k):
                    e[k] = data[k]
        return set_status("rejected", extra)

    @app.post("/api/queue/ignore")
    def api_ignore():
        return set_status("ignored")

    @app.get("/api/errors")
    def api_errors():
        items = [e for e in queue.load() if e.get("status") == "error"]
        return jsonify(sorted(items, key=_sort_key, reverse=True))

    @app.post("/api/errors/dismiss")
    def api_errors_dismiss():
        return set_status("dismissed")

    @app.post("/api/errors/retry")
    def api_errors_retry():
        return set_status("approved")

    @app.get("/api/tasks")
    def api_tasks():
        items = [e for e in queue.load() if e.get("category") == "task" and e.get("status") == "open_task"]
        return jsonify(sorted(items, key=_sort_key, reverse=True))

    @app.post("/api/tasks/done")
    def api_tasks_done():
        return set_status("task_done")

    @app.get("/api/corrections")
    def api_corrections():
        return jsonify(read_json(base_dir / "corrections.json", default=[]) or [])

    @app.post("/api/email/open")
    def api_email_open():
        try:
            if not open_mail:
                raise RuntimeError("Öffnen im Mailprogramm ist nicht verfügbar")
            open_mail(cfg(), body().get("entry_id") or "")
            return jsonify(success=True)
        except Exception as exc:
            return jsonify(success=False, error=str(exc))

    # ── Bot steuern ───────────────────────────────────────────────────
    @app.post("/api/bot/start")
    def api_bot_start():
        if runner:
            runner.resume()
        return jsonify(success=True)

    @app.post("/api/bot/stop")
    def api_bot_stop():
        if runner:
            runner.pause()
        return jsonify(success=True)

    # ── Konfiguration ─────────────────────────────────────────────────
    @app.route("/api/config", methods=["GET", "POST"])
    def api_config():
        path = base_dir / "config.json"
        if request.method == "GET":
            data = read_json(path, default={})
            return Response(json.dumps(data, ensure_ascii=False, indent=2), mimetype="application/json")
        data = json.loads(request.get_data(as_text=True))  # ungültiges JSON -> 500 mit Fehlermeldung
        if not isinstance(data, dict):
            raise ValueError("Konfiguration muss ein JSON-Objekt sein")
        write_json(path, data)
        log.info("Konfiguration ueber das Dashboard gespeichert")
        return jsonify(success=True)

    # ── Oberfläche ────────────────────────────────────────────────────
    @app.get("/")
    def index():
        return send_from_directory(www, "index.html")

    @app.get("/<path:filename>")
    def static_file(filename: str):
        return send_from_directory(www, filename)

    @app.errorhandler(Exception)
    def on_error(exc):
        from werkzeug.exceptions import HTTPException
        if isinstance(exc, HTTPException):
            return exc
        return jsonify(error=str(exc)), 500

    return app
