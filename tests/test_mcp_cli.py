"""Tests for the optional MCP command-line entry points."""

from __future__ import annotations

import subprocess
import sys


def _run_module(module: str, option: str) -> str:
    result = subprocess.run(
        [sys.executable, "-m", module, option],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr
    return result.stdout


def test_mcp_server_version() -> None:
    output = _run_module("easy_uiauto.mcp.server", "--version")
    assert output.strip() == "easy_uiauto 0.1.13"


def test_mcp_service_help() -> None:
    output = _run_module("easy_uiauto.mcp.service", "--help")
    assert "easy_uiauto_service" in output
    assert "--version" in output


def test_mcp_server_help_lists_client_configuration_actions() -> None:
    output = _run_module("easy_uiauto.mcp.server", "--help")
    assert "--install-codex" in output
    assert "--uninstall-codex" in output
    assert "--install-claude-code" in output
    assert "--uninstall-claude-code" in output
