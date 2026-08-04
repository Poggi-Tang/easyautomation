"""Tests for the optional MCP command-line entry points."""

from __future__ import annotations

import subprocess
import sys


def _run_module(module: str, option: str) -> str:
    result = subprocess.run(
        [sys.executable, "-m", module, option],
        check=True,
        capture_output=True,
        text=True,
    )
    return result.stdout


def test_mcp_server_version() -> None:
    output = _run_module("easy_uiauto.mcp.server", "--version")
    assert output.strip() == "easy_uiauto 0.1.9"


def test_mcp_service_help() -> None:
    output = _run_module("easy_uiauto.mcp.service", "--help")
    assert "easy_uiauto_service" in output
    assert "--version" in output
