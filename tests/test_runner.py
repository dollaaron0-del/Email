"""Bot-Thread: Start/Pause aus dem Dashboard, Verbindung erst wenn der Bot läuft."""
import time

from emailbot.runner import BotRunner

from .conftest import FakeBackend, make_mail


def wait_for(cond, timeout=5.0):
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        if cond():
            return True
        time.sleep(0.05)
    return False


def test_paused_runner_connects_only_after_resume(base_dir):
    backend = FakeBackend()
    backend.add(make_mail("Rechnung MOCK_SONSTIGES"))
    calls = []

    def factory(cfg):
        calls.append(cfg["mail_backend"])
        return backend

    runner = BotRunner(base_dir, factory, start_paused=True)
    runner.start()
    try:
        time.sleep(0.5)
        assert calls == [] and runner.running is False  # angehalten: kein Mailzugriff
        runner.resume()
        assert wait_for(lambda: backend.moves)  # verbunden, gescannt, Mail eingereiht
        assert calls == ["outlook"] and runner.running is True
        runner.pause()
        assert runner.running is False
    finally:
        runner.stop()


def test_unimplemented_backend_pauses_instead_of_crashing(base_dir):
    def factory(cfg):
        raise NotImplementedError("Graph noch nicht umgesetzt")

    runner = BotRunner(base_dir, factory)
    runner.start()
    try:
        assert wait_for(lambda: not runner.running)
        assert runner._thread.is_alive()  # Dashboard bleibt nutzbar
    finally:
        runner.stop()
