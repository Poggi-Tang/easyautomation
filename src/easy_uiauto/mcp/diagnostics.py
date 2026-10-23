"""Installation checks and end-to-end diagnostics for easy_uiauto."""

from __future__ import annotations

import base64
import importlib.util
import io
import json
import math
import os
import re
import shutil
import subprocess
import sys
import time
from collections.abc import Callable, Mapping
from pathlib import Path
from urllib import error as urlerror
from urllib import request as urlrequest
from urllib.parse import urlsplit

TESSERACT_PACKAGE_ID = "UB-Mannheim.TesseractOCR"

_VISION_SETTINGS = (
    "EASY_UIAUTO_VISION_API_URL",
    "EASY_UIAUTO_VISION_API_KEY",
    "EASY_UIAUTO_VISION_MODEL",
)

# These probes never capture the screen, send input, install packages, or use the
# network. Run native imports and UIA calls out of process so a broken provider
# cannot hang the MCP server. Only the exit status is exposed to callers.
_READINESS_PROBES = {
    "uia": """
import uiautomation
root = uiautomation.GetRootControl()
assert root is not None
root.GetChildren()
""",
    "image_matching": """
import cv2
import numpy
from PIL import Image
template = numpy.array([[0, 20], [90, 255]], dtype=numpy.uint8)
canvas = numpy.zeros((6, 6), dtype=numpy.uint8)
canvas[2:4, 3:5] = template
matches = cv2.matchTemplate(canvas, template, cv2.TM_SQDIFF)
assert cv2.minMaxLoc(matches)[2] == (3, 2)
Image.fromarray(canvas)
""",
    "ocr": """
import pytesseract
from PIL import Image
Image.new('RGB', (1, 1))
""",
}


def _modules_missing(modules: tuple[str, ...]) -> list[str]:
    missing = []
    for module in modules:
        try:
            found = importlib.util.find_spec(module) is not None
        except (ImportError, ValueError):
            found = False
        if not found:
            missing.append(module)
    return missing


def _readiness_probe(name: str, timeout: float) -> str | None:
    """Return a stable error code, never native stderr or configuration values."""
    try:
        result = subprocess.run(
            [sys.executable, "-c", _READINESS_PROBES[name]],
            # Inherited MCP stdin can block Python startup on Windows while the
            # transport reader is waiting on the same pipe.
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return "probe_timeout"
    except OSError:
        return "probe_start_failed"
    return None if result.returncode == 0 else "local_probe_failed"


def _ocr_languages(timeout: float) -> tuple[set[str], str | None]:
    """Use the same command as pytesseract, without changing PATH or its config."""
    loaded = sys.modules.get("pytesseract.pytesseract")
    command = getattr(loaded, "tesseract_cmd", "tesseract")
    if not shutil.which(command):
        return set(), "tesseract_executable_missing"
    try:
        result = subprocess.run(
            [command, "--list-langs"],
            stdin=subprocess.DEVNULL,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return set(), "probe_timeout"
    except OSError:
        return set(), "tesseract_start_failed"
    if result.returncode != 0:
        return set(), "tesseract_probe_failed"
    # Tesseract versions differ in whether --list-langs writes to stdout or
    # stderr. Only language identifiers may enter the public report.
    languages = {
        line.strip()
        for line in (result.stdout + "\n" + result.stderr).splitlines()
        if re.fullmatch(r"[A-Za-z0-9_]+(?:/[A-Za-z0-9_]+)*", line.strip())
    }
    return languages, None


def get_capability_readiness(
    language: str = "eng",
    probe_timeout: float = 3.0,
    *,
    vision_settings: Mapping[str, str] | None = None,
) -> dict:
    """Report local prerequisites separately from complete tool verification.

    ``available`` describes each capability's stated ``scope``. OCR and image
    matching checks cover local engines, not desktop capture. Remote vision is
    ``available=None`` when configured because connectivity and authentication
    are deliberately not tested. No screen contents or secret values are read
    into the report. Each subprocess is bounded by ``probe_timeout`` seconds.

    MCP callers should supply vision settings from configuration's dynamic
    resolver, which also understands Windows user and existing Codex bindings.
    Standalone callers default to the current process environment.
    """
    if not math.isfinite(probe_timeout) or not 0 < probe_timeout <= 30:
        raise ValueError("probe_timeout must be greater than 0 and no greater than 30")
    language_pattern = r"[A-Za-z0-9_]+(?:/[A-Za-z0-9_]+)*(?:\+[A-Za-z0-9_]+(?:/[A-Za-z0-9_]+)*)*"
    if not re.fullmatch(language_pattern, language):
        raise ValueError("language must contain Tesseract language codes, e.g. eng+chi_sim")

    capabilities = {}
    requirements = {
        "uia": ("uiautomation",),
        "ocr": ("pytesseract", "PIL"),
        "image_matching": ("cv2", "numpy", "PIL"),
    }
    scopes = {
        "uia": "Desktop root access; individual controls and actions are not tested.",
        "ocr": (
            "OCR imports, executable and requested languages; "
            "recognition and screen capture are not tested."
        ),
        "image_matching": "Synthetic local image matching; screen capture is not tested.",
    }
    for name, modules in requirements.items():
        missing = _modules_missing(modules)
        if name == "uia" and sys.platform != "win32":
            missing.append("Windows")
        status = {
            "configured": not missing,
            "available": False,
            "runtime_verified": False,
            "missing": missing,
            "issues": [],
            "scope": scopes[name],
        }
        if not missing:
            issue = _readiness_probe(name, probe_timeout)
            if issue:
                status["issues"].append(issue)
            else:
                status["available"] = True
                status["runtime_verified"] = name == "uia"
        if name == "ocr":
            languages, issue = _ocr_languages(probe_timeout)
            status["requested_languages"] = sorted(set(language.split("+")))
            status["installed_languages"] = sorted(languages)
            if issue:
                status["issues"].append(issue)
                status["available"] = False
                status["configured"] = False
            else:
                missing_languages = sorted(set(language.split("+")) - languages)
                status["missing"].extend(f"language:{item}" for item in missing_languages)
                if missing_languages:
                    status["configured"] = status["available"] = False
        status["status"] = "ready" if status["available"] else "unavailable"
        capabilities[name] = status

    settings = os.environ if vision_settings is None else vision_settings
    missing_settings = [name for name in _VISION_SETTINGS if not settings.get(name, "").strip()]
    remote_issues = []
    api_url = settings.get(_VISION_SETTINGS[0], "").strip()
    if api_url:
        try:
            parsed_url = urlsplit(api_url)
            valid_url = parsed_url.scheme in {"http", "https"} and bool(parsed_url.hostname)
        except ValueError:
            valid_url = False
        if not valid_url:
            remote_issues.append("invalid_api_url")
    configured = not missing_settings and not remote_issues
    capabilities["remote_vision"] = {
        "configured": configured,
        "available": None if configured else False,
        "runtime_verified": False,
        "status": "unverified" if configured else "unavailable",
        "missing": missing_settings,
        "issues": remote_issues,
        "scope": "Configuration presence and URL syntax only; no network request is made.",
    }
    return {
        "platform": sys.platform,
        "network_checked": False,
        "screen_capture_checked": False,
        "capabilities": capabilities,
    }


def _timed_step(name: str, operation: Callable[[], str]) -> dict:
    started = time.perf_counter()
    try:
        detail = operation()
        return {
            "name": name,
            "ok": True,
            "timing_ms": round((time.perf_counter() - started) * 1000, 1),
            "detail": detail,
        }
    except Exception as error:
        return {
            "name": name,
            "ok": False,
            "timing_ms": round((time.perf_counter() - started) * 1000, 1),
            "detail": f"{type(error).__name__}: {error}",
        }


def ensure_python_vision_dependencies(version: str) -> dict:
    """Install the current package's vision extra only when modules are missing."""

    def ensure() -> str:
        missing = [
            module
            for module in ("cv2", "pytesseract")
            if importlib.util.find_spec(module) is None
        ]
        if not missing:
            return "OpenCV and pytesseract are already installed"
        result = subprocess.run(
            [
                sys.executable,
                "-m",
                "pip",
                "install",
                "--upgrade",
                f"easy-uiauto[mcp,vision]=={version}",
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode != 0:
            detail = result.stderr.strip() or result.stdout.strip() or "unknown pip error"
            raise RuntimeError(f"Could not install the vision extra: {detail}")
        return f"Installed missing Python vision modules: {', '.join(missing)}"

    return _timed_step("Python vision dependencies", ensure)


def _tesseract_candidates() -> list[Path]:
    candidates = []
    command = shutil.which("tesseract")
    if command:
        candidates.append(Path(command))
    for root_name in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        root = os.environ.get(root_name)
        if not root:
            continue
        candidates.extend(
            [
                Path(root) / "Tesseract-OCR" / "tesseract.exe",
                Path(root) / "Programs" / "Tesseract-OCR" / "tesseract.exe",
            ]
        )
    return candidates


def _find_tesseract() -> Path | None:
    for candidate in _tesseract_candidates():
        if candidate.is_file():
            return candidate.resolve()
    return None


def _add_to_user_path(directory: Path) -> None:
    """Make a discovered Tesseract executable available to future MCP processes."""
    directory_text = str(directory)
    current_parts = [part for part in os.environ.get("PATH", "").split(os.pathsep) if part]
    if not any(part.casefold() == directory_text.casefold() for part in current_parts):
        os.environ["PATH"] = directory_text + os.pathsep + os.environ.get("PATH", "")
    if os.name != "nt":
        return

    import winreg

    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
        try:
            current, value_type = winreg.QueryValueEx(key, "Path")
        except FileNotFoundError:
            current, value_type = "", winreg.REG_EXPAND_SZ
        user_parts = [part for part in str(current).split(os.pathsep) if part]
        if any(part.casefold() == directory_text.casefold() for part in user_parts):
            return
        updated = os.pathsep.join([*user_parts, directory_text])
        winreg.SetValueEx(key, "Path", 0, value_type, updated)


def ensure_tesseract() -> tuple[dict, str]:
    """Find or install Tesseract and return its executable path."""
    located_path = ""

    def ensure() -> str:
        nonlocal located_path
        executable = _find_tesseract()
        installed = False
        if executable is None:
            winget = shutil.which("winget")
            if winget is None:
                raise RuntimeError(
                    "Tesseract is missing and winget was not found; install Tesseract OCR"
                )
            result = subprocess.run(
                [
                    winget,
                    "install",
                    "--id",
                    TESSERACT_PACKAGE_ID,
                    "--exact",
                    "--silent",
                    "--accept-package-agreements",
                    "--accept-source-agreements",
                    "--disable-interactivity",
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            if result.returncode != 0:
                detail = result.stderr.strip() or result.stdout.strip() or "unknown winget error"
                raise RuntimeError(f"Could not install Tesseract with winget: {detail}")
            executable = _find_tesseract()
            installed = True
        if executable is None:
            raise RuntimeError("Tesseract installation completed but tesseract.exe was not found")
        _add_to_user_path(executable.parent)
        located_path = str(executable)
        action = "Installed" if installed else "Found"
        return f"{action} Tesseract at {executable}"

    result = _timed_step("Tesseract", ensure)
    return result, located_path


def _test_uia() -> str:
    import uiautomation

    root = uiautomation.GetRootControl()
    if root is None:
        raise RuntimeError("UI Automation did not return the desktop root")
    children = root.GetChildren()
    return f"Desktop root accessible; discovered {len(children)} top-level controls"


def _test_ocr(tesseract_path: str) -> str:
    import pytesseract
    from PIL import Image, ImageDraw, ImageFont

    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
    image = Image.new("RGB", (720, 150), "white")
    draw = ImageDraw.Draw(image)
    font_path = Path(os.environ.get("WINDIR", "C:\\Windows")) / "Fonts" / "arial.ttf"
    font = (
        ImageFont.truetype(str(font_path), 48)
        if font_path.is_file()
        else ImageFont.load_default()
    )
    expected = "EASY UIAUTO 123"
    draw.text((24, 38), expected, fill="black", font=font)
    detected = pytesseract.image_to_string(image, lang="eng", config="--psm 7").strip()
    normalized = re.sub(r"[^A-Z0-9]", "", detected.upper())
    if "EASYUIAUTO123" not in normalized:
        raise RuntimeError(f"OCR output did not match the synthetic test text: {detected!r}")
    return f"Recognized synthetic text: {expected}"


def _json_content(content: str) -> dict:
    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*|\s*```$", "", content, flags=re.IGNORECASE)
    result = json.loads(content)
    if not isinstance(result, dict):
        raise RuntimeError("Remote vision response was not a JSON object")
    return result


def _test_remote_vision(api_url: str, api_key: str, model: str, version: str) -> str:
    from PIL import Image, ImageDraw, ImageFont

    image = Image.new("RGB", (640, 240), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((160, 70, 480, 180), fill="#1769aa")
    font_path = Path(os.environ.get("WINDIR", "C:\\Windows")) / "Fonts" / "arial.ttf"
    font = (
        ImageFont.truetype(str(font_path), 36)
        if font_path.is_file()
        else ImageFont.load_default()
    )
    draw.text((220, 102), "TEST TARGET", fill="white", font=font)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    image_url = "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")
    payload = {
        "model": model,
        "temperature": 0,
        "messages": [
            {
                "role": "system",
                "content": "Inspect the supplied image and return JSON only.",
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": (
                            "Does the image contain a blue rectangle labelled TEST TARGET? "
                            'Return exactly {"visible":true} or {"visible":false}.'
                        ),
                    },
                    {"type": "image_url", "image_url": {"url": image_url}},
                ],
            },
        ],
    }
    if model.lower().startswith("deepseek-"):
        payload["thinking"] = {"type": "disabled"}
    request = urlrequest.Request(
        api_url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "User-Agent": f"easy-uiauto/{version}",
        },
        method="POST",
    )
    try:
        with urlrequest.urlopen(request, timeout=45) as response:
            body = json.loads(response.read().decode("utf-8"))
    except urlerror.HTTPError as error:
        raise RuntimeError(f"Remote vision request failed with HTTP {error.code}") from error
    except urlerror.URLError as error:
        raise RuntimeError(f"Remote vision request failed: {error.reason}") from error
    try:
        content = body["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as error:
        raise RuntimeError("Remote vision response did not contain a chat completion") from error
    if isinstance(content, list):
        content = "".join(item.get("text", "") for item in content if isinstance(item, dict))
    result = _json_content(str(content))
    if result.get("visible") is not True:
        raise RuntimeError("Remote model did not identify the synthetic test target")
    return f"Remote model {model} identified the synthetic visual target"


def run_full_diagnostics(
    api_url: str,
    api_key: str,
    model: str,
    version: str,
    tesseract_path: str,
) -> list[dict]:
    """Run deterministic UIA, local OCR, and remote vision checks."""
    return [
        _timed_step("UIA", _test_uia),
        _timed_step("OCR", lambda: _test_ocr(tesseract_path)),
        _timed_step(
            "Remote AI vision",
            lambda: _test_remote_vision(api_url, api_key, model, version),
        ),
    ]


def format_report(steps: list[dict]) -> str:
    """Format setup and diagnostic results without exposing credentials."""
    lines = []
    for step in steps:
        state = "PASS" if step["ok"] else "FAIL"
        lines.append(
            f"[{state}] {step['name']}: {step['timing_ms']:.1f} ms - {step['detail']}"
        )
    return "\n".join(lines)
