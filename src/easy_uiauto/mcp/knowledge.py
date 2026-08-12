"""Obsidian-compatible UI knowledge vault.

Markdown, YAML frontmatter, and PNG files are the source of truth. The JSON
index is disposable and can always be rebuilt from the Markdown records.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

KNOWLEDGE_DIR = "EASY_UIAUTO_KNOWLEDGE_DIR"
INDEX_VERSION = 1
SEARCHABLE_STATUSES = {"verified", "observed", "suspect"}
EXECUTABLE_STATUSES = {"verified"}


def vault_root() -> Path:
    configured = os.environ.get(KNOWLEDGE_DIR, "").strip()
    return Path(configured).expanduser() if configured else Path.home() / "easy_uiauto_vault"


def slugify(value: str, fallback: str = "item") -> str:
    value = re.sub(r"\s+", "-", str(value).strip().lower())
    value = re.sub(r"[^\w.\-\u4e00-\u9fff]+", "-", value, flags=re.UNICODE)
    value = re.sub(r"-+", "-", value).strip("-._")
    return value[:80] or fallback


def stable_id(prefix: str, *parts: Any) -> str:
    encoded = json.dumps(parts, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
    return f"{slugify(prefix)}-{hashlib.sha256(encoded).hexdigest()[:12]}"


def utc_now() -> str:
    return datetime.now(UTC).isoformat(timespec="seconds")


def app_dir(app_id: str, root: Path | None = None) -> Path:
    return (root or vault_root()) / "applications" / slugify(app_id, "application")


def initialize_app(app_id: str, app_name: str, root: Path | None = None) -> Path:
    directory = app_dir(app_id, root)
    for child in (
        "pages",
        "regions",
        "controls",
        "operations",
        "images/pages",
        "images/controls",
        "quarantine",
        ".easy_uiauto",
    ):
        (directory / child).mkdir(parents=True, exist_ok=True)
    app_record = {
        "kind": "application",
        "id": slugify(app_id, "application"),
        "name": app_name,
        "updated_at": utc_now(),
        "schema_version": INDEX_VERSION,
    }
    write_markdown(directory / "app.md", app_record, f"# {app_name}\n\nUI knowledge root.\n")
    return directory


def _frontmatter_value(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def write_markdown(path: Path, metadata: dict, body: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = ["---"]
    for key in sorted(metadata):
        lines.append(f"{key}: {_frontmatter_value(metadata[key])}")
    lines.extend(["---", "", body.rstrip(), ""])
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text("\n".join(lines), encoding="utf-8")
    temporary.replace(path)


def read_markdown(path: Path) -> tuple[dict, str]:
    text = path.read_text(encoding="utf-8")
    if not text.startswith("---\n"):
        raise ValueError(f"Missing YAML frontmatter: {path}")
    frontmatter, separator, body = text[4:].partition("\n---\n")
    if not separator:
        raise ValueError(f"Unterminated YAML frontmatter: {path}")
    metadata = {}
    for line in frontmatter.splitlines():
        key, marker, raw_value = line.partition(":")
        if not marker:
            continue
        try:
            metadata[key.strip()] = json.loads(raw_value.strip())
        except json.JSONDecodeError:
            metadata[key.strip()] = raw_value.strip()
    return metadata, body.lstrip("\n")


def control_path(directory: Path, control_id: str, status: str) -> Path:
    folder = "quarantine" if status == "quarantined" else "controls"
    return directory / folder / f"{slugify(control_id)}.md"


def _control_body(record: dict) -> str:
    title = record.get("semantic_name") or record.get("name") or record["id"]
    actions = record.get("actions", [])
    action_lines = "\n".join(f"- `{action}`" for action in actions) or "- Observation only"
    return (
        f"# {title}\n\n"
        f"Application: [[../app|{record.get('app_name', record['app_id'])}]]\n\n"
        f"Page: `{record.get('page_id', 'unknown')}`  \n"
        f"Region: `{record.get('region_id', 'unassigned')}`  \n"
        f"Status: **{record.get('status', 'observed')}**\n\n"
        "## Operations\n\n"
        f"{action_lines}\n\n"
        "## Notes\n\n"
        f"{record.get('notes', '')}\n"
    )


def save_control(directory: Path, record: dict) -> Path:
    existing = find_control_record(directory, record["id"], include_quarantine=True)
    if existing:
        old_path, old_record = existing
        record.setdefault("created_at", old_record.get("created_at", utc_now()))
        expected_folder = "quarantine" if record["status"] == "quarantined" else "controls"
        if old_path.parent.name != expected_folder:
            old_path.unlink(missing_ok=True)
    else:
        record.setdefault("created_at", utc_now())
    record["updated_at"] = utc_now()
    record["kind"] = "control"
    path = control_path(directory, record["id"], record.get("status", "observed"))
    write_markdown(path, record, _control_body(record))
    return path


def save_page(directory: Path, record: dict) -> Path:
    record = {**record, "kind": "page", "updated_at": utc_now()}
    title = record.get("name") or record["id"]
    body = f"# {title}\n\n{record.get('description', '')}\n"
    path = directory / "pages" / f"{slugify(record['id'])}.md"
    write_markdown(path, record, body)
    return path


def save_region(directory: Path, record: dict) -> Path:
    record = {**record, "kind": "region", "updated_at": utc_now()}
    title = record.get("name") or record["id"]
    body = f"# {title}\n\n{record.get('description', '')}\n"
    path = directory / "regions" / f"{slugify(record['id'])}.md"
    write_markdown(path, record, body)
    return path


def iter_records(directory: Path, kind: str | None = None) -> list[tuple[Path, dict]]:
    patterns = {
        "control": ["controls/*.md", "quarantine/*.md"],
        "page": ["pages/*.md"],
        "region": ["regions/*.md"],
        "application": ["app.md"],
    }
    selected = patterns.get(kind, [pattern for values in patterns.values() for pattern in values])
    records = []
    for pattern in selected:
        for path in sorted(directory.glob(pattern)):
            try:
                metadata, _body = read_markdown(path)
            except (OSError, ValueError):
                continue
            records.append((path, metadata))
    return records


def find_control_record(
    directory: Path,
    control_id: str,
    include_quarantine: bool = False,
) -> tuple[Path, dict] | None:
    folders = ["controls", "quarantine"] if include_quarantine else ["controls"]
    filename = f"{slugify(control_id)}.md"
    for folder in folders:
        path = directory / folder / filename
        if path.is_file():
            metadata, _body = read_markdown(path)
            return path, metadata
    return None


def rebuild_index(directory: Path) -> dict:
    app_record = {}
    app_path = directory / "app.md"
    if app_path.is_file():
        app_record, _body = read_markdown(app_path)
    pages = [record for _path, record in iter_records(directory, "page")]
    regions = [record for _path, record in iter_records(directory, "region")]
    controls = [record for _path, record in iter_records(directory, "control")]
    index = {
        "schema_version": INDEX_VERSION,
        "generated_at": utc_now(),
        "application": app_record,
        "pages": pages,
        "regions": regions,
        "controls": controls,
    }
    index_path = directory / ".easy_uiauto" / "index.json"
    index_path.parent.mkdir(parents=True, exist_ok=True)
    index_path.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    return index


def load_index(directory: Path, rebuild: bool = False) -> dict:
    path = directory / ".easy_uiauto" / "index.json"
    if not rebuild and path.is_file():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            pass
    return rebuild_index(directory)


def list_apps(root: Path | None = None) -> list[dict]:
    applications = (root or vault_root()) / "applications"
    result = []
    if not applications.is_dir():
        return result
    for directory in sorted(path for path in applications.iterdir() if path.is_dir()):
        index = load_index(directory)
        result.append(
            {
                "id": index.get("application", {}).get("id", directory.name),
                "name": index.get("application", {}).get("name", directory.name),
                "controls": len(index.get("controls", [])),
                "verified": sum(
                    control.get("status") == "verified" for control in index.get("controls", [])
                ),
                "quarantined": sum(
                    control.get("status") == "quarantined"
                    for control in index.get("controls", [])
                ),
            }
        )
    return result


def _search_text(record: dict) -> str:
    values = [
        record.get("id", ""),
        record.get("semantic_name", ""),
        record.get("name", ""),
        record.get("control_type", ""),
        record.get("automation_id", ""),
        record.get("page_id", ""),
        record.get("region_id", ""),
        record.get("notes", ""),
        " ".join(record.get("tags", [])),
        " ".join(record.get("actions", [])),
    ]
    return " ".join(str(value) for value in values).casefold()


def search_controls(
    directory: Path,
    query: str = "",
    statuses: set[str] | None = None,
    limit: int = 50,
) -> list[dict]:
    words = [word for word in query.casefold().split() if word]
    statuses = statuses or SEARCHABLE_STATUSES
    controls = load_index(directory).get("controls", [])
    matches = []
    for control in controls:
        if control.get("status") not in statuses:
            continue
        haystack = _search_text(control)
        if words and not all(word in haystack for word in words):
            continue
        matches.append(control)
    matches.sort(
        key=lambda item: (
            item.get("status") != "verified",
            not item.get("is_key", False),
            item.get("command", item.get("id", "")),
        )
    )
    return matches[: max(1, limit)]


def available_commands(directory: Path, page_id: str = "") -> list[dict]:
    commands = []
    for control in search_controls(directory, statuses=EXECUTABLE_STATUSES, limit=10000):
        if page_id and control.get("page_id") != page_id:
            continue
        for action in control.get("actions", []):
            commands.append(
                {
                    "command": f"{control['command']}.{action}",
                    "control_id": control["id"],
                    "semantic_name": control.get("semantic_name") or control.get("name"),
                    "action": action,
                    "page": control.get("page_id"),
                    "region": control.get("region_id"),
                    "status": control.get("status"),
                }
            )
    return commands


def resolve_command(directory: Path, command: str) -> tuple[dict, str]:
    for item in available_commands(directory):
        if item["command"] == command:
            found = find_control_record(directory, item["control_id"])
            if found is None:
                break
            return found[1], item["action"]
    raise KeyError(f"Unknown or unverified UI command: {command}")


def replace_image(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_suffix(destination.suffix + ".tmp")
    shutil.copyfile(source, temporary)
    temporary.replace(destination)


def write_command_catalog(directory: Path) -> Path:
    """Write a human-readable UI CLI catalog derived from verified controls."""
    commands = available_commands(directory)
    lines = ["# UI Command Catalog", "", "Generated from verified control records.", ""]
    current_page = None
    for item in commands:
        if item["page"] != current_page:
            current_page = item["page"]
            lines.extend([f"## {current_page}", ""])
        lines.append(
            f"- `{item['command']}` - {item.get('semantic_name') or item['control_id']} "
            f"({item.get('region') or 'unassigned'})"
        )
    if not commands:
        lines.append("No verified commands are available.")
    path = directory / "operations" / "UI-CLI.md"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path
