"""Install the easy_uiauto Codex skills bundled with the Python package."""

from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path

SKILL_NAMES = ("easy-uiauto-learning", "easy-uiauto-operate")


def bundled_skills_dir() -> Path:
    return Path(__file__).resolve().parent.parent / "skills"


def codex_skills_dir() -> Path:
    codex_home = Path(os.environ.get("CODEX_HOME", Path.home() / ".codex"))
    return codex_home / "skills"


def install_codex_skills() -> str:
    """Install or update only the two easy_uiauto skills."""
    source_root = bundled_skills_dir()
    destination_root = codex_skills_dir()
    destination_root.mkdir(parents=True, exist_ok=True)
    installed = []
    for name in SKILL_NAMES:
        source = source_root / name
        if not (source / "SKILL.md").is_file():
            raise RuntimeError(f"Bundled skill is missing: {source}")
        temporary_parent = Path(tempfile.mkdtemp(prefix=f".{name}-", dir=destination_root))
        temporary = temporary_parent / name
        try:
            shutil.copytree(source, temporary)
            destination = destination_root / name
            backup = destination_root / f".{name}.backup"
            if backup.exists():
                shutil.rmtree(backup)
            if destination.exists():
                destination.replace(backup)
            temporary.replace(destination)
            if backup.exists():
                shutil.rmtree(backup)
        finally:
            if temporary_parent.exists():
                shutil.rmtree(temporary_parent)
        installed.append(str(destination))
    return "Installed Codex skills:\n" + "\n".join(f"  {path}" for path in installed)


def uninstall_codex_skills() -> str:
    """Remove only the easy_uiauto skills from the current Codex home."""
    removed = []
    for name in SKILL_NAMES:
        destination = codex_skills_dir() / name
        if destination.exists():
            shutil.rmtree(destination)
            removed.append(str(destination))
    if not removed:
        return "No easy_uiauto Codex skills were installed."
    return "Removed Codex skills:\n" + "\n".join(f"  {path}" for path in removed)
