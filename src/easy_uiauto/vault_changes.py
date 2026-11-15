"""Small per-application revisions for cross-process library refresh."""
from __future__ import annotations

from pathlib import Path
import time
from uuid import uuid4


def notify(path):
    """Publish after a record commit; the marker contains no user data."""
    path = Path(path)
    directory = next((p for p in (path, *path.parents) if p.parent.name == "applications"), None)
    if directory is None:return
    from .mcp.image_storage import _safe_path
    marker = directory / ".easy_uiauto" / "revision"
    if _safe_path(directory, marker) is None:raise ValueError("知识库变更标记路径无效")
    marker.parent.mkdir(parents=True, exist_ok=True)
    temporary = marker.with_name("revision." + uuid4().hex + ".tmp")
    try:
        temporary.write_text(uuid4().hex, encoding="ascii")
        for attempt in range(6):
            try:
                temporary.replace(marker)
                break
            except PermissionError:
                if attempt == 5:raise
                time.sleep(.01 * (attempt + 1))
    finally:temporary.unlink(missing_ok=True)


def revisions(root):
    """Read one tiny token per application, never walk controls or image trees."""
    from .mcp.image_storage import _safe_path
    result = {}
    for directory in (Path(root) / "applications").glob("*"):
        if _safe_path(Path(root), directory) is None or not (directory / "app.md").is_file():continue
        marker = directory / ".easy_uiauto" / "revision"
        if _safe_path(directory, marker) is None:continue
        try:result[directory.name] = marker.read_text(encoding="ascii")
        except FileNotFoundError:result[directory.name] = str((directory / "app.md").stat().st_mtime_ns)
    return result


def revision(directory):
    marker = Path(directory) / ".easy_uiauto" / "revision"
    try:return marker.read_text(encoding="ascii")
    except FileNotFoundError:return ""
