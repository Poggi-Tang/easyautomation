"""Tests for full Codex setup diagnostics."""

from __future__ import annotations

from collections import deque
from pathlib import Path
from types import SimpleNamespace

from easy_uiauto.mcp import diagnostics


def test_python_vision_dependencies_are_not_reinstalled_when_present(monkeypatch) -> None:
    monkeypatch.setattr(diagnostics.importlib.util, "find_spec", lambda _module: object())
    monkeypatch.setattr(
        diagnostics.subprocess,
        "run",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("pip was not expected")
        ),
    )

    result = diagnostics.ensure_python_vision_dependencies("1.2.3")

    assert result["ok"] is True
    assert "already installed" in result["detail"]


def test_existing_tesseract_is_added_to_user_path(monkeypatch, tmp_path) -> None:
    executable = tmp_path / "Tesseract-OCR" / "tesseract.exe"
    calls: list[Path] = []
    monkeypatch.setattr(diagnostics, "_find_tesseract", lambda: executable)
    monkeypatch.setattr(diagnostics, "_add_to_user_path", calls.append)

    result, path = diagnostics.ensure_tesseract()

    assert result["ok"] is True
    assert path == str(executable)
    assert calls == [executable.parent]


def test_missing_tesseract_is_installed_with_winget(monkeypatch, tmp_path) -> None:
    executable = tmp_path / "Tesseract-OCR" / "tesseract.exe"
    discoveries = deque([None, executable])
    commands: list[list[str]] = []
    monkeypatch.setattr(diagnostics, "_find_tesseract", discoveries.popleft)
    monkeypatch.setattr(
        diagnostics.shutil,
        "which",
        lambda name: "winget.exe" if name == "winget" else None,
    )
    monkeypatch.setattr(
        diagnostics.subprocess,
        "run",
        lambda command, **_kwargs: commands.append(command)
        or SimpleNamespace(returncode=0, stdout="", stderr=""),
    )
    monkeypatch.setattr(diagnostics, "_add_to_user_path", lambda _path: None)

    result, path = diagnostics.ensure_tesseract()

    assert result["ok"] is True
    assert path == str(executable)
    assert commands == [
        [
            "winget.exe",
            "install",
            "--id",
            diagnostics.TESSERACT_PACKAGE_ID,
            "--exact",
            "--silent",
            "--accept-package-agreements",
            "--accept-source-agreements",
            "--disable-interactivity",
        ]
    ]


def test_report_formats_pass_and_failure() -> None:
    report = diagnostics.format_report(
        [
            {"name": "UIA", "ok": True, "timing_ms": 1.2, "detail": "ready"},
            {"name": "OCR", "ok": False, "timing_ms": 3.4, "detail": "failed"},
        ]
    )

    assert "[PASS] UIA: 1.2 ms - ready" in report
    assert "[FAIL] OCR: 3.4 ms - failed" in report
