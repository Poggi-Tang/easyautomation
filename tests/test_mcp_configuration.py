"""Tests for Codex and Claude Code MCP configuration commands."""

from __future__ import annotations

import sys

from easy_uiauto.mcp import configuration, diagnostics, skill_installation


def test_reads_only_existing_easy_uiauto_mcp_environment(monkeypatch, tmp_path) -> None:
    codex_home = tmp_path / ".codex"
    codex_home.mkdir()
    (codex_home / "config.toml").write_text(
        """
[mcp_servers.easy_uiauto]
command = "python"

[mcp_servers.easy_uiauto.env]
EASY_UIAUTO_VISION_API_KEY = "existing-secret"

[mcp_servers.unrelated.env]
EASY_UIAUTO_VISION_API_KEY = "wrong-secret"
""".strip(),
        encoding="utf-8",
    )
    monkeypatch.setenv("CODEX_HOME", str(codex_home))

    assert configuration._read_existing_codex_environment(
        configuration.VISION_API_KEY
    ) == "existing-secret"


def test_install_codex_uses_current_python(monkeypatch) -> None:
    calls: list[tuple[str, list[str]]] = []

    def fake_run(name: str, arguments: list[str]) -> str:
        calls.append((name, arguments))
        return "added"

    monkeypatch.setattr(configuration, "_run_client", fake_run)

    assert configuration.install_codex() == "added"
    assert calls == [
        (
            "codex",
            [
                "mcp",
                "add",
                "easy_uiauto",
                "--",
                sys.executable,
                "-m",
                "easy_uiauto.mcp.server",
            ],
        )
    ]


def test_install_claude_code_uses_user_scope(monkeypatch) -> None:
    calls: list[tuple[str, list[str]]] = []

    def fake_run(name: str, arguments: list[str]) -> str:
        calls.append((name, arguments))
        return "added"

    monkeypatch.setattr(configuration, "_run_client", fake_run)

    assert configuration.install_claude_code() == "added"
    assert calls == [
        (
            "claude",
            [
                "mcp",
                "add",
                "--scope",
                "user",
                "easy_uiauto",
                "--",
                sys.executable,
                "-m",
                "easy_uiauto.mcp.server",
            ],
        )
    ]


def test_quick_setup_codex_reuses_existing_key_and_replaces_entry(monkeypatch) -> None:
    writes: list[tuple[str, str]] = []
    actions: list[str] = []

    monkeypatch.setattr(
        configuration,
        "_existing_vision_value",
        lambda name: "secret" if name == configuration.VISION_API_KEY else "",
    )
    monkeypatch.setattr(
        configuration,
        "_write_user_environment",
        lambda name, value: writes.append((name, value)),
    )
    monkeypatch.setattr(configuration, "uninstall_codex", lambda: actions.append("remove"))
    monkeypatch.setattr(configuration, "install_codex", lambda: actions.append("add"))
    monkeypatch.setattr(configuration, "show_codex", lambda: "configured entry")
    monkeypatch.setattr(
        configuration,
        "_prompt_api_key",
        lambda: (_ for _ in ()).throw(AssertionError("key prompt was not expected")),
    )

    output = configuration.quick_setup_codex(
        "https://api.example/v1/chat/completions",
        "vision-model",
        "1.2.3",
    )

    assert writes == [
        (configuration.VISION_API_URL, "https://api.example/v1/chat/completions"),
        (configuration.VISION_API_KEY, "secret"),
        (configuration.VISION_MODEL, "vision-model"),
    ]
    assert actions == ["remove", "add"]
    assert "easy_uiauto 1.2.3" in output
    assert "secret" not in output
    assert "Restart Codex once" in output
    assert "available to a current easy_uiauto MCP process immediately" in output


def test_quick_setup_codex_prompts_once_for_missing_key(monkeypatch) -> None:
    prompts: list[str] = []

    monkeypatch.setattr(configuration, "_existing_vision_value", lambda _name: "")
    monkeypatch.setattr(configuration, "_write_user_environment", lambda _name, _value: None)
    monkeypatch.setattr(configuration, "uninstall_codex", lambda: "")
    monkeypatch.setattr(configuration, "install_codex", lambda: "")
    monkeypatch.setattr(configuration, "show_codex", lambda: "configured entry")
    monkeypatch.setattr(
        configuration,
        "_prompt_api_key",
        lambda: prompts.append("key") or "new-secret",
    )

    output = configuration.quick_setup_codex(
        "https://api.example/v1/chat/completions",
        "vision-model",
        "1.2.3",
    )

    assert prompts == ["key"]
    assert "new-secret" not in output


def test_quick_setup_codex_rejects_invalid_url_before_writing(monkeypatch) -> None:
    monkeypatch.setattr(
        configuration,
        "_write_user_environment",
        lambda _name, _value: (_ for _ in ()).throw(AssertionError("write was not expected")),
    )

    try:
        configuration.quick_setup_codex("not-a-url", "vision-model", "1.2.3")
    except RuntimeError as error:
        assert "must start with" in str(error)
    else:
        raise AssertionError("invalid URL was accepted")


def test_full_setup_codex_runs_all_validation_steps(monkeypatch) -> None:
    monkeypatch.setattr(
        configuration,
        "_configure_remote_vision",
        lambda _url, _model: ("https://api.example/v1", "secret", "vision-model"),
    )
    monkeypatch.setattr(configuration, "_replace_codex", lambda: "configured entry")
    monkeypatch.setattr(
        skill_installation,
        "install_codex_skills",
        lambda: "Installed Codex skills: test",
    )
    monkeypatch.setattr(
        diagnostics,
        "ensure_python_vision_dependencies",
        lambda _version: {
            "name": "Python vision dependencies",
            "ok": True,
            "timing_ms": 1.0,
            "detail": "ready",
        },
    )
    monkeypatch.setattr(
        diagnostics,
        "ensure_tesseract",
        lambda: (
            {"name": "Tesseract", "ok": True, "timing_ms": 2.0, "detail": "ready"},
            "tesseract.exe",
        ),
    )
    monkeypatch.setattr(
        diagnostics,
        "run_full_diagnostics",
        lambda **_kwargs: [
            {"name": "UIA", "ok": True, "timing_ms": 3.0, "detail": "ready"},
            {"name": "OCR", "ok": True, "timing_ms": 4.0, "detail": "ready"},
            {
                "name": "Remote AI vision",
                "ok": True,
                "timing_ms": 5.0,
                "detail": "ready",
            },
        ],
    )

    output = configuration.full_setup_codex("", "", "1.2.3")

    assert "[PASS] UIA" in output
    assert "[PASS] OCR" in output
    assert "[PASS] Remote AI vision" in output
    assert "secret" not in output
    assert "Restart Codex once" in output
    assert "do not restart again just for environment changes" in output


def test_vision_configuration_status_never_returns_key(monkeypatch) -> None:
    values = {
        configuration.VISION_API_URL: "https://api.example/v1/chat/completions",
        configuration.VISION_API_KEY: "secret-value",
        configuration.VISION_MODEL: "vision-model",
    }
    monkeypatch.setattr(
        configuration,
        "_existing_vision_value",
        lambda name: values.get(name, ""),
    )

    status = configuration.vision_configuration_status()

    assert status["ready"] is True
    assert status["api_key_configured"] is True
    assert "secret-value" not in str(status)


def test_full_setup_codex_fails_when_a_validation_fails(monkeypatch) -> None:
    monkeypatch.setattr(
        configuration,
        "_configure_remote_vision",
        lambda _url, _model: ("https://api.example/v1", "secret", "vision-model"),
    )
    monkeypatch.setattr(configuration, "_replace_codex", lambda: "configured entry")
    monkeypatch.setattr(
        skill_installation,
        "install_codex_skills",
        lambda: "Installed Codex skills: test",
    )
    monkeypatch.setattr(
        diagnostics,
        "ensure_python_vision_dependencies",
        lambda _version: {
            "name": "Python vision dependencies",
            "ok": True,
            "timing_ms": 1.0,
            "detail": "ready",
        },
    )
    monkeypatch.setattr(
        diagnostics,
        "ensure_tesseract",
        lambda: (
            {"name": "Tesseract", "ok": True, "timing_ms": 2.0, "detail": "ready"},
            "tesseract.exe",
        ),
    )
    monkeypatch.setattr(
        diagnostics,
        "run_full_diagnostics",
        lambda **_kwargs: [
            {"name": "OCR", "ok": False, "timing_ms": 3.0, "detail": "failed"}
        ],
    )

    try:
        configuration.full_setup_codex("", "", "1.2.3")
    except RuntimeError as error:
        assert "[FAIL] OCR" in str(error)
    else:
        raise AssertionError("failed validation was accepted")
