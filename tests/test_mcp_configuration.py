"""Tests for Codex and Claude Code MCP configuration commands."""

from __future__ import annotations

import sys

from easy_uiauto.mcp import configuration


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
