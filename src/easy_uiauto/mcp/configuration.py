"""MCP client configuration helpers for the easy_uiauto command line."""

from __future__ import annotations

import getpass
import os
import shutil
import subprocess
import sys
from pathlib import Path

import tomllib

SERVER_NAME = "easy_uiauto"
SERVER_MODULE = "easy_uiauto.mcp.server"
VISION_API_URL = "EASY_UIAUTO_VISION_API_URL"
VISION_API_KEY = "EASY_UIAUTO_VISION_API_KEY"
VISION_MODEL = "EASY_UIAUTO_VISION_MODEL"


def _read_user_environment(name: str) -> str:
    """Read a user-scoped environment variable without exposing it."""
    if os.name != "nt":
        return os.environ.get(name, "")
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
            value, _value_type = winreg.QueryValueEx(key, name)
            return str(value)
    except FileNotFoundError:
        return ""


def _write_user_environment(name: str, value: str) -> None:
    """Persist one user-scoped environment variable without using setx."""
    if not value:
        raise ValueError(f"{name} must not be empty")
    if os.name != "nt":
        raise RuntimeError("Quick Codex setup is supported only on Windows")
    import winreg

    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
    os.environ[name] = value


def _read_existing_codex_environment(name: str) -> str:
    """Read only easy_uiauto's existing MCP environment, never Codex auth data."""
    codex_home = Path(os.environ.get("CODEX_HOME", Path.home() / ".codex"))
    config_path = codex_home / "config.toml"
    try:
        with config_path.open("rb") as config_file:
            config = tomllib.load(config_file)
    except (FileNotFoundError, OSError, tomllib.TOMLDecodeError):
        return ""
    server = config.get("mcp_servers", {}).get(SERVER_NAME, {})
    value = server.get("env", {}).get(name, "")
    return str(value) if value else ""


def _existing_vision_value(name: str) -> str:
    """Resolve a vision setting without consulting unrelated credentials."""
    return (
        _read_user_environment(name).strip()
        or os.environ.get(name, "").strip()
        or _read_existing_codex_environment(name).strip()
    )


def _prompt_api_key() -> str:
    """Read the API key from a terminal or a hidden Windows dialog."""
    if sys.stdin.isatty():
        return getpass.getpass("Vision API key (input hidden): ").strip()
    if os.name != "nt":
        raise RuntimeError(
            "No interactive terminal is available; set EASY_UIAUTO_VISION_API_KEY first"
        )

    import tkinter
    from tkinter import simpledialog

    root = tkinter.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        value = simpledialog.askstring(
            "easy_uiauto quick setup",
            "Vision API key:",
            show="*",
            parent=root,
        )
        return (value or "").strip()
    finally:
        root.destroy()


def _prompt_required_text(label: str, option: str) -> str:
    """Prompt for a non-secret value or explain how to provide it non-interactively."""
    try:
        return input(f"{label}: ").strip()
    except EOFError as error:
        raise RuntimeError(f"{label} is required; pass {option}") from error


def _configure_remote_vision(api_url: str, model: str) -> tuple[str, str, str]:
    """Resolve and persist remote vision settings without exposing the API key."""
    api_url = api_url.strip() or _existing_vision_value(VISION_API_URL)
    model = model.strip() or _existing_vision_value(VISION_MODEL)
    if not api_url:
        api_url = _prompt_required_text("Vision API URL", "--vision-url URL")
    if not model:
        model = _prompt_required_text("Vision model", "--vision-model MODEL")
    if not api_url.startswith(("https://", "http://")):
        raise RuntimeError("Vision API URL must start with https:// or http://")
    if not model:
        raise RuntimeError("Vision model must not be empty")

    api_key = _existing_vision_value(VISION_API_KEY)
    if not api_key:
        api_key = _prompt_api_key()
    if not api_key:
        raise RuntimeError("Vision API key must not be empty")

    _write_user_environment(VISION_API_URL, api_url)
    _write_user_environment(VISION_API_KEY, api_key)
    _write_user_environment(VISION_MODEL, model)
    return api_url, api_key, model


def _replace_codex() -> str:
    """Replace only the global easy_uiauto Codex MCP entry."""
    try:
        uninstall_codex()
    except RuntimeError:
        pass
    install_codex()
    return show_codex()


def quick_setup_codex(api_url: str, model: str, version: str) -> str:
    """Configure remote vision and replace the global easy_uiauto Codex entry."""
    api_url, _api_key, model = _configure_remote_vision(api_url, model)
    configured = _replace_codex()

    return (
        f"easy_uiauto {version}\n"
        f"Vision API URL: {api_url}\n"
        f"Vision model: {model}\n"
        "Vision API key: configured in the Windows user environment (value hidden)\n"
        f"{configured}\n"
        "Quick setup complete. Fully restart Codex to load the MCP server and environment."
    )


def full_setup_codex(api_url: str, model: str, version: str) -> str:
    """Install required components, configure Codex, and verify every backend."""
    from . import diagnostics

    api_url, api_key, model = _configure_remote_vision(api_url, model)
    steps = [diagnostics.ensure_python_vision_dependencies(version)]
    tesseract_result, tesseract_path = diagnostics.ensure_tesseract()
    steps.append(tesseract_result)

    configured = _replace_codex()
    if all(step["ok"] for step in steps):
        steps.extend(
            diagnostics.run_full_diagnostics(
                api_url=api_url,
                api_key=api_key,
                model=model,
                version=version,
                tesseract_path=tesseract_path,
            )
        )
    report = diagnostics.format_report(steps)
    if not all(step["ok"] for step in steps):
        raise RuntimeError(f"Full setup validation failed:\n{report}")

    return (
        f"easy_uiauto {version}\n"
        f"Vision API URL: {api_url}\n"
        f"Vision model: {model}\n"
        "Vision API key: configured in the Windows user environment (value hidden)\n"
        f"{report}\n"
        f"{configured}\n"
        "Full setup and validation complete. Fully restart Codex to load the MCP server "
        "and environment."
    )


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
