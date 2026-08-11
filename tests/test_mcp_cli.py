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
    assert output.strip() == "easy_uiauto 0.1.21"


def test_mcp_service_help() -> None:
    output = _run_module("easy_uiauto.mcp.service", "--help")
    assert "easy_uiauto_service" in output
    assert "--version" in output


def test_mcp_server_help_lists_client_configuration_actions() -> None:
    output = _run_module("easy_uiauto.mcp.server", "--help")
    assert "--install-codex" in output
    assert "--quick-setup-codex" in output
    assert "--full-setup-codex" in output
    assert "--vision-url" in output
    assert "--vision-model" in output
    assert "--uninstall-codex" in output
    assert "--install-claude-code" in output
    assert "--uninstall-claude-code" in output


def test_mcp_server_help_explains_core_workflows() -> None:
    output = _run_module("easy_uiauto.mcp.server", "--help")

    assert "Control location workflow (preferred):" in output
    assert "find_control(location=LOCATION)" in output
    assert "run_record(write_file=True)" in output
    assert "from easy_uiauto.record import run_record" in output
    assert "run_action(action_json='<recorded action JSON>')" in output
    assert "UIA LOCATION -> OCR or image template -> remote AI vision" in output
    assert "EASY_UIAUTO_VISION_API_URL" in output
    assert "Restart the client" in output


def test_mcp_server_help_explains_testing_and_diagnostics() -> None:
    output = _run_module("easy_uiauto.mcp.server", "--help")

    assert "Testing and diagnostics:" in output
    assert "Read-only MCP smoke test:" in output
    assert "python -m pytest -q" in output
    assert "find_text_on_screen(text=..., language='eng')" in output
    assert "find_control_by_vision(description=...)" in output
    assert "sends the selected screenshot" in output
