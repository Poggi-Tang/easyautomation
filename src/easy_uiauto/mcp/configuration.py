"""MCP client configuration helpers for the easy_uiauto command line."""

from __future__ import annotations

import getpass
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from urllib import error as urlerror, parse as urlparse, request as urlrequest

import tomllib

SERVER_NAME = "easy_uiauto"
SERVER_MODULE = "easy_uiauto.mcp.server"
VISION_API_URL = "EASY_UIAUTO_VISION_API_URL"
VISION_API_KEY = "EASY_UIAUTO_VISION_API_KEY"
VISION_MODEL = "EASY_UIAUTO_VISION_MODEL"
KNOWLEDGE_DIR = "EASY_UIAUTO_KNOWLEDGE_DIR"


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


def knowledge_directory() -> Path:
    """Read the saved directory on each call, including in an already running MCP."""
    configured = _existing_vision_value(KNOWLEDGE_DIR)
    return Path(configured).expanduser() if configured else Path.home() / "easy_uiauto_vault"


def load_gui_configuration() -> dict[str, str]:
    """Read only this application's settings for the local configuration window."""
    return {
        "knowledge_dir": str(knowledge_directory()),
        "api_url": _existing_vision_value(VISION_API_URL),
        "api_key": _existing_vision_value(VISION_API_KEY),
        "api_model": _existing_vision_value(VISION_MODEL),
    }


def api_endpoints(api_url: str) -> tuple[str, str]:
    """Accept a provider base, versioned base, models URL or chat-completions URL."""
    value = api_url.strip()
    try:
        parts = urlparse.urlsplit(value)
        valid = parts.scheme in {"http", "https"} and parts.hostname
        _ = parts.port  # Validate an explicitly supplied port as well.
    except ValueError as error:
        raise ValueError("API URL 格式不正确") from error
    if not valid or parts.username or parts.password or parts.query or parts.fragment:
        raise ValueError("API URL 须为 http:// 或 https:// 地址，不含账号、查询参数或片段")
    if any(character.isspace() for character in value):
        raise ValueError("API URL 不能包含空格或换行")
    path = parts.path.rstrip("/")
    for suffix in ("/chat/completions", "/models"):
        if path.endswith(suffix):
            path = path[:-len(suffix)]
            break
    base = urlparse.urlunsplit((parts.scheme, parts.netloc, path, "", ""))
    return base + "/chat/completions", base + "/models"


def fetch_models(api_url: str, api_key: str, timeout: float = 20) -> list[str]:
    """Fetch an OpenAI-compatible model catalogue; never include credentials in errors."""
    _chat_url, models_url = api_endpoints(api_url)
    api_key = api_key.strip()
    if not api_key or "\r" in api_key or "\n" in api_key:
        raise ValueError("请填写有效的 API Key")
    request = urlrequest.Request(
        models_url,
        headers={"Authorization": f"Bearer {api_key}", "Accept": "application/json"},
        method="GET",
    )
    try:
        with urlrequest.urlopen(request, timeout=timeout) as response:
            body = response.read(2 * 1024 * 1024 + 1)
        if len(body) > 2 * 1024 * 1024:
            raise ValueError("模型列表响应过大，请手动填写模型名称")
        result = json.loads(body.decode("utf-8-sig"))
    except urlerror.HTTPError as error:
        hints = {
            401: "请检查 API Key", 403: "当前 API Key 没有访问权限",
            404: "该地址没有模型列表接口，可手动填写 API Model",
            405: "接口不支持获取模型列表，可手动填写 API Model",
            429: "请求过于频繁，请稍后再试",
        }
        detail = f"获取失败（HTTP {error.code}）：" + hints.get(error.code, "请检查服务状态")
        error.close()
        raise RuntimeError(detail) from None
    except (TimeoutError, urlerror.URLError) as error:
        reason = getattr(error, "reason", error)
        detail = "请求超时" if isinstance(reason, TimeoutError) else "无法连接服务"
        raise RuntimeError(detail + "，请检查 API URL 或稍后重试") from None
    except (UnicodeError, json.JSONDecodeError):
        raise RuntimeError("接口未返回有效 JSON，可手动填写 API Model") from None
    if not isinstance(result, dict) or not isinstance(result.get("data"), list):
        raise RuntimeError("接口未返回兼容的模型列表，可手动填写 API Model")
    models = sorted({
        item["id"].strip() for item in result["data"]
        if isinstance(item, dict) and isinstance(item.get("id"), str) and item["id"].strip()
    }, key=str.casefold)
    if not models:
        raise RuntimeError("接口返回的模型列表为空，可手动填写 API Model")
    return models


def save_gui_configuration(knowledge_dir: str, api_url: str, api_key: str,
                           api_model: str) -> dict[str, str]:
    """Validate the entire form before persisting any of its four values."""
    if os.name != "nt":
        raise RuntimeError("配置窗口的持久化保存仅支持 Windows")
    if not knowledge_dir.strip():
        raise ValueError("请选择或填写知识库路径")
    directory = Path(os.path.expandvars(knowledge_dir.strip())).expanduser()
    if not directory.is_absolute():
        raise ValueError("知识库路径须为绝对路径")
    chat_url, _models_url = api_endpoints(api_url)
    api_key, api_model = api_key.strip(), api_model.strip()
    if not api_key or "\r" in api_key or "\n" in api_key:
        raise ValueError("请填写有效的 API Key")
    if not api_model:
        raise ValueError("请输入或选择 API Model")
    directory.mkdir(parents=True, exist_ok=True)
    values = {
        KNOWLEDGE_DIR: str(directory), VISION_API_URL: chat_url,
        VISION_API_KEY: api_key, VISION_MODEL: api_model,
    }
    import winreg

    # Roll back a partial write. The process environment is updated only on success.
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
        previous = {}
        written = []
        try:
            for name, value in values.items():
                try:
                    previous[name] = winreg.QueryValueEx(key, name)
                except FileNotFoundError:
                    previous[name] = None
                winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
                written.append(name)
        except OSError:
            for name in reversed(written):
                old_value = previous[name]
                if old_value is None:
                    winreg.DeleteValue(key, name)
                else:
                    winreg.SetValueEx(key, name, 0, old_value[1], old_value[0])
            raise
    os.environ.update(values)
    # Let future processes inherit the changes. A running MCP reads the registry itself.
    import ctypes
    from ctypes import wintypes

    result = ctypes.c_size_t()
    send = ctypes.WinDLL("user32", use_last_error=True).SendMessageTimeoutW
    send.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPCWSTR,
                     wintypes.UINT, wintypes.UINT, ctypes.POINTER(ctypes.c_size_t)]
    send.restype = wintypes.LPARAM
    send(0xFFFF, 0x001A, 0, "Environment", 0x0002, 1000, ctypes.byref(result))
    return {"knowledge_dir": str(directory), "api_url": chat_url, "api_model": api_model}


def vision_request_options(model: str) -> dict:
    """Avoid DeepSeek's default extended thinking for interactive UI requests."""
    if model.lower().startswith("deepseek-"):
        return {"thinking": {"type": "disabled"}}
    return {}


def vision_configuration_status() -> dict:
    """Return non-secret readiness details using values visible to current calls."""
    api_url = _existing_vision_value(VISION_API_URL)
    api_key = _existing_vision_value(VISION_API_KEY)
    model = _existing_vision_value(VISION_MODEL)
    ready = bool(api_url and api_key and model)
    return {
        "ready": ready,
        "api_url_configured": bool(api_url),
        "api_key_configured": bool(api_key),
        "model_configured": bool(model),
        "model": model,
        "detail": (
            "UI learning configuration is ready."
            if ready
            else (
                "Run easy_uiauto --quick-setup-codex --vision-url URL "
                "--vision-model MODEL. Configuration changes are read dynamically; "
                "do not retry with full-uia."
            )
        ),
    }


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


def replace_codex_preserving_settings() -> str:
    """Move the MCP entry to this Python while retaining this app's settings."""
    for name in (KNOWLEDGE_DIR, VISION_API_URL, VISION_API_KEY, VISION_MODEL):
        value = _existing_vision_value(name)
        if value:
            _write_user_environment(name, value)
    return _replace_codex()


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
        "Quick setup complete. Vision settings are available to a current easy_uiauto MCP "
        "process immediately. Restart Codex once only if the MCP command or package changed."
    )


def full_setup_codex(api_url: str, model: str, version: str) -> str:
    """Install required components, configure Codex, and verify every backend."""
    from . import diagnostics

    api_url, api_key, model = _configure_remote_vision(api_url, model)
    steps = [diagnostics.ensure_python_vision_dependencies(version)]
    tesseract_result, tesseract_path = diagnostics.ensure_tesseract()
    steps.append(tesseract_result)

    configured = _replace_codex()
    from .skill_installation import install_codex_skills

    skills = install_codex_skills()
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
        f"{skills}\n"
        "Full setup and validation complete. Vision settings are available to a current "
        "easy_uiauto MCP process immediately. Restart Codex once to load a newly installed "
        "MCP package or updated skills; do not restart again just for environment changes."
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
