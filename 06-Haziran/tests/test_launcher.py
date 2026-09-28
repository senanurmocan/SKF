"""Launcher path and network binding regression tests."""

import sys
from pathlib import Path
from types import SimpleNamespace

import run

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
import run_app


def test_root_launcher_delegates_to_project_launcher(monkeypatch):
    captured = {}

    def fake_run(command, *, cwd, check):
        captured.update(command=command, cwd=cwd, check=check)
        return SimpleNamespace(returncode=7)

    monkeypatch.setattr(run_app.subprocess, "run", fake_run)
    monkeypatch.setattr(sys, "argv", ["run_app.py", "--port", "8512"])

    assert run_app.main() == 7
    assert Path(captured["command"][1]) == Path(run_app.__file__).resolve().parent / "06-Haziran" / "run.py"
    assert captured["command"][-2:] == ["--port", "8512"]
    assert Path(captured["cwd"]) == Path(run_app.__file__).resolve().parent / "06-Haziran"
    assert captured["check"] is False


def _capture_streamlit_command(monkeypatch, arguments):
    captured = {}
    monkeypatch.setattr(run, "find_free_port", lambda _start: 8505)
    monkeypatch.setattr(
        run.subprocess,
        "run",
        lambda command, **kwargs: captured.update(command=command, **kwargs)
        or SimpleNamespace(returncode=0),
    )
    monkeypatch.setattr(sys, "argv", ["run.py", *arguments])
    run.main()
    return captured["command"]


def _option(command, name):
    return command[command.index(name) + 1]


def test_launcher_defaults_to_localhost(monkeypatch):
    command = _capture_streamlit_command(monkeypatch, [])

    assert _option(command, "--server.address") == "127.0.0.1"
    assert _option(command, "--server.enableCORS") == "false"
    assert _option(command, "--browser.gatherUsageStats") == "false"


def test_lan_mode_enables_xsrf_protection(monkeypatch):
    command = _capture_streamlit_command(monkeypatch, ["--lan"])

    assert _option(command, "--server.address") == "0.0.0.0"
    assert _option(command, "--server.enableXsrfProtection") == "true"


def test_lan_open_is_explicit_and_disables_xsrf(monkeypatch):
    command = _capture_streamlit_command(monkeypatch, ["--lan-open"])

    assert _option(command, "--server.address") == "0.0.0.0"
    assert _option(command, "--server.enableXsrfProtection") == "false"
