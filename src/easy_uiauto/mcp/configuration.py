"""MCP client configuration helpers for the easy_uiauto command line."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys


SERVER_NAME = "easy_uiauto"
SERVER_MODULE = "easy_uiauto.mcp.server"


def _client_executable(name: str) -> str:
    """Return a runnable client executable, including Windows ``.cmd`` shims."""
    if os.name == "nt":
        command = shutil.which(f"{name}.cmd") or shutil.which(name)
    else:
        command = shutil.which(name)
    if command is None:
        raise RuntimeError(f"{name} command was not found on PATH")
    return command


def _run_client(name: str, arguments: list[str]) -> str:
    result = subprocess.run(
        [_client_executable(name), *arguments],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.returncode != 0:
        details = result.stderr.strip() or result.stdout.strip() or "unknown error"
        raise RuntimeError(f"{name} MCP command failed: {details}")
    return (result.stdout or "").strip()


def install_codex() -> str:
    """Register easy_uiauto globally through the Codex CLI."""
    return _run_client(
        "codex",
        ["mcp", "add", SERVER_NAME, "--", sys.executable, "-m", SERVER_MODULE],
    )


def uninstall_codex() -> str:
    """Remove the global easy_uiauto entry through the Codex CLI."""
    return _run_client("codex", ["mcp", "remove", SERVER_NAME])


def show_codex() -> str:
    """Return the configured Codex MCP entry."""
    return _run_client("codex", ["mcp", "get", SERVER_NAME])


def install_claude_code() -> str:
    """Register easy_uiauto at Claude Code's user scope."""
    return _run_client(
        "claude",
        [
            "mcp",
            "add",
            "--scope",
            "user",
            SERVER_NAME,
            "--",
            sys.executable,
            "-m",
            SERVER_MODULE,
        ],
    )


def uninstall_claude_code() -> str:
    """Remove the user-scoped easy_uiauto entry from Claude Code."""
    return _run_client("claude", ["mcp", "remove", "--scope", "user", SERVER_NAME])


def show_claude_code() -> str:
    """Return the configured Claude Code MCP entry."""
    return _run_client("claude", ["mcp", "get", SERVER_NAME])
