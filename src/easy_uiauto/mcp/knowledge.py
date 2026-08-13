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
INDEX_VERSION = 3
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


def normalized_rect(rectangle: dict, reference: dict, relative: bool = False) -> dict:
    """Convert volatile pixels to a window-relative visual hint in the 0..1 range."""
    reference_width = reference.get("width")
    if reference_width is None:
        reference_width = reference["right"] - reference["left"]
    reference_height = reference.get("height")
    if reference_height is None:
        reference_height = reference["bottom"] - reference["top"]
    width = max(1, int(reference_width))
    height = max(1, int(reference_height))
    origin_left = 0 if relative else int(reference.get("left", 0))
    origin_top = 0 if relative else int(reference.get("top", 0))
    return {
        "left": round((int(rectangle["left"]) - origin_left) / width, 6),
        "top": round((int(rectangle["top"]) - origin_top) / height, 6),
        "right": round((int(rectangle["right"]) - origin_left) / width, 6),
        "bottom": round((int(rectangle["bottom"]) - origin_top) / height, 6),
    }


def _public_window_observation(item: dict) -> dict:
    return {
        key: item.get(key)
        for key in ("title", "class_name", "control_type", "foreground")
        if key in item
    }


def sanitize_interaction_geometry(record: dict) -> dict:
    """Remove process-local handles, PIDs, and absolute desktop positions before persistence."""
    record = dict(record)
    observed_size = None
    for snapshot_name in ("before", "after"):
        snapshot = dict(record.get(snapshot_name, {}))
        rectangle = snapshot.pop("window_rect", None)
        if isinstance(rectangle, dict):
            size = {
                "width": rectangle.get("width", 0),
                "height": rectangle.get("height", 0),
            }
            snapshot["observed_window_size"] = size
            observed_size = observed_size or size
        for key in ("target_handle", "desktop_origin"):
            snapshot.pop(key, None)
        snapshot["windows"] = [
            _public_window_observation(item)
            for item in snapshot.get("windows", [])
            if isinstance(item, dict)
        ]
        action_control = dict(snapshot.get("action_control", {}))
        action_control.pop("rect", None)
        snapshot["action_control"] = action_control
        record[snapshot_name] = snapshot

    changes = dict(record.get("window_changes", {}))
    for key in ("added", "removed"):
        changes[key] = [
            _public_window_observation(item)
            for item in changes.get(key, [])
            if isinstance(item, dict)
        ]
    changes["changed"] = [
        {
            "before": _public_window_observation(item.get("before", {})),
            "after": _public_window_observation(item.get("after", {})),
        }
        for item in changes.get("changed", [])
        if isinstance(item, dict)
    ]
    record["window_changes"] = changes
    record["transient_windows"] = [
        _public_window_observation(item)
        for item in record.get("transient_windows", [])
        if isinstance(item, dict)
    ]
    for key in ("changed_controls", "popup_controls"):
        values = []
        for item in record.get(key, []):
            if not isinstance(item, dict):
                continue
            cleaned = dict(item)
            cleaned.pop("rect", None)
            cleaned.pop("process_id", None)
            cleaned.pop("handle", None)
            values.append(cleaned)
        record[key] = values
    if observed_size and record.get("changed_region_geometry_role") != "visual-hint-only":
        for key in ("changed_regions",):
            record[key] = [
                normalized_rect(item, observed_size, relative=True)
                for item in record.get(key, [])
                if isinstance(item, dict)
            ]
        record["changed_region_geometry_role"] = "visual-hint-only"
    action_changes = dict(record.get("action_property_changes", {}))
    if "rect" in action_changes:
        action_changes["geometry_changed"] = True
        action_changes.pop("rect", None)
    record["action_property_changes"] = action_changes
    stability = dict(record.get("stability", {}))
    if observed_size and stability.get("dynamic_region_geometry_role") != "visual-hint-only":
        stability["dynamic_regions"] = [
            normalized_rect(item, observed_size, relative=True)
            for item in stability.get("dynamic_regions", [])
            if isinstance(item, dict)
        ]
        stability["dynamic_region_geometry_role"] = "visual-hint-only"
    stability["observed_windows"] = [
        _public_window_observation(item)
        for item in stability.get("observed_windows", [])
        if isinstance(item, dict)
    ]
    record["stability"] = stability
    return record


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
        "images/interactions",
        "interactions",
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
        "## Meaning\n\n"
        f"{record.get('description') or 'No semantic description.'}\n\n"
        f"Intent: `{record.get('intent', 'unknown')}`  \n"
        f"Confidence: `{record.get('semantic_confidence', 0)}`  \n"
        f"Source: `{record.get('semantic_source', 'unknown')}`  \n"
        f"Risk: `{record.get('risk', 'unknown')}`  \n"
        f"Function verification: "
        f"`{record.get('function_verification', {}).get('status', 'unknown')}`\n\n"
        "## Operations\n\n"
        f"{action_lines}\n\n"
        "## Notes\n\n"
        f"{record.get('notes', '')}\n"
    )


def save_control(directory: Path, record: dict) -> Path:
    record = dict(record)
    record.pop("rect", None)
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


def teach_control(
    directory: Path,
    control_id: str,
    semantic_name: str,
    intent: str,
    description: str,
    actions: list[str] | None = None,
    aliases: list[str] | None = None,
    risk: str = "safe",
    requires_confirmation: bool = False,
) -> dict:
    """Apply a human-confirmed meaning without weakening locator verification."""
    found = find_control_record(directory, control_id, include_quarantine=True)
    if found is None:
        raise KeyError(f"Unknown control: {control_id}")
    _path, record = found
    semantic_name = semantic_name.strip()
    normalized_intent = slugify(intent, "")
    if not semantic_name or not normalized_intent or not description.strip():
        raise ValueError("semantic_name, intent, and description are required")
    supported = record.get("supported_actions", record.get("actions", []))
    selected = actions if actions is not None else supported
    selected = list(dict.fromkeys(action for action in selected if action in supported))
    if supported and not selected:
        raise ValueError(f"actions must include one of: {', '.join(supported)}")
    risk = risk.strip().lower()
    if risk not in {"safe", "state-changing", "external", "destructive", "unknown"}:
        raise ValueError(f"Unsupported risk: {risk}")
    positioning_ok = record.get("verification", {}).get("image") == "passed" and (
        record.get("verification", {}).get("location") == "passed"
        or record.get("visual_fallback_ready") is True
    )
    record.update(
        {
            "semantic_name": semantic_name,
            "intent": normalized_intent,
            "description": description.strip(),
            "actions": selected,
            "aliases": list(
                dict.fromkeys(value.strip() for value in (aliases or []) if value.strip())
            ),
            "risk": risk,
            "requires_confirmation": bool(requires_confirmation)
            or risk in {"external", "destructive"},
            "semantic_confidence": 1.0,
            "semantic_status": "manual",
            "semantic_source": "manual",
            "semantic_evidence": ["human teaching"],
            "semantic_ambiguity": "",
            "function_verification": {
                "status": "human-confirmed",
                "method": "manual teaching",
                "executed": record.get("function_verification", {}).get("executed", False),
                "verified_at": utc_now(),
            },
            "status": "verified" if positioning_ok else "quarantined",
        }
    )
    command = ".".join(
        (
            slugify(record.get("page_id", "main")),
            slugify(record.get("region_id", "unassigned")),
            normalized_intent,
        )
    )
    occupied = {
        item.get("command")
        for _item_path, item in iter_records(directory, "control")
        if item.get("id") != control_id
    }
    record["command"] = f"{command}-{slugify(control_id)[-6:]}" if command in occupied else command
    record["tags"] = list(
        dict.fromkeys(
            [
                *record.get("tags", []),
                normalized_intent,
                *(aliases or []),
            ]
        )
    )
    save_control(directory, record)
    rebuild_index(directory)
    write_command_catalog(directory)
    return record


def save_page(directory: Path, record: dict) -> Path:
    record = {**record, "kind": "page", "updated_at": utc_now()}
    record.pop("rect", None)
    title = record.get("name") or record["id"]
    body = f"# {title}\n\n{record.get('description', '')}\n"
    path = directory / "pages" / f"{slugify(record['id'])}.md"
    write_markdown(path, record, body)
    return path


def save_region(directory: Path, record: dict) -> Path:
    record = {**record, "kind": "region", "updated_at": utc_now()}
    record.pop("rect", None)
    title = record.get("name") or record["id"]
    body = f"# {title}\n\n{record.get('description', '')}\n"
    path = directory / "regions" / f"{slugify(record['id'])}.md"
    write_markdown(path, record, body)
    return path


def save_interaction(directory: Path, record: dict) -> Path:
    """Persist one before/after operation effect as Markdown source of truth."""
    record = sanitize_interaction_geometry(
        {**record, "kind": "interaction", "updated_at": utc_now()}
    )
    title = record.get("semantic_name") or record.get("command") or record["id"]
    effects = record.get("effects", [])
    effect_lines = (
        "\n".join(
            f"- `{item.get('type', 'change')}`: {item.get('description', '')}"
            for item in effects
            if isinstance(item, dict)
        )
        or "- No confirmed effect"
    )
    body = (
        f"# {title}\n\n"
        f"Command: `{record.get('command', '')}`  \n"
        f"Status: **{record.get('status', 'unknown')}**  \n"
        f"Before state: `{record.get('before_state_id', '')}`  \n"
        f"After state: `{record.get('after_state_id', '')}`\n\n"
        "## Effects\n\n"
        f"{effect_lines}\n\n"
        "## Success Condition\n\n"
        f"{record.get('success_condition', 'Not established')}\n\n"
        "## Recovery\n\n"
        f"{record.get('recovery', {}).get('detail', 'Not attempted')}\n"
    )
    path = directory / "interactions" / f"{slugify(record['id'])}.md"
    write_markdown(path, record, body)
    return path


def iter_records(directory: Path, kind: str | None = None) -> list[tuple[Path, dict]]:
    patterns = {
        "control": ["controls/*.md", "quarantine/*.md"],
        "page": ["pages/*.md"],
        "region": ["regions/*.md"],
        "application": ["app.md"],
        "interaction": ["interactions/*.md"],
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


def migrate_legacy_volatile_geometry(directory: Path) -> int:
    """Upgrade v2 absolute geometry and process-local interaction fields in place."""
    changed = 0
    page_rectangles = {}
    page_records = iter_records(directory, "page")
    for _path, page in page_records:
        if isinstance(page.get("rect"), dict):
            page_rectangles[page.get("id")] = page["rect"]

    for path, record in iter_records(directory, "control"):
        rectangle = record.get("rect")
        reference = page_rectangles.get(record.get("page_id"))
        if not isinstance(rectangle, dict) or not isinstance(reference, dict):
            continue
        _metadata, body = read_markdown(path)
        record["normalized_rect"] = normalized_rect(rectangle, reference)
        record["geometry_role"] = "visual-hint-only"
        record.pop("rect", None)
        write_markdown(path, record, body)
        changed += 1

    for path, record in iter_records(directory, "region"):
        rectangle = record.get("rect")
        reference = page_rectangles.get(record.get("page_id"))
        if not isinstance(rectangle, dict) or not isinstance(reference, dict):
            continue
        _metadata, body = read_markdown(path)
        record["normalized_rect"] = normalized_rect(rectangle, reference, relative=True)
        record["geometry_role"] = "visual-hint-only"
        record.pop("rect", None)
        write_markdown(path, record, body)
        changed += 1

    for path, record in page_records:
        if "rect" not in record:
            continue
        _metadata, body = read_markdown(path)
        record.pop("rect", None)
        write_markdown(path, record, body)
        changed += 1

    for path, record in iter_records(directory, "interaction"):
        cleaned = sanitize_interaction_geometry(record)
        if cleaned == record:
            continue
        _metadata, body = read_markdown(path)
        write_markdown(path, cleaned, body)
        changed += 1
    return changed


def rebuild_index(directory: Path) -> dict:
    migrate_legacy_volatile_geometry(directory)
    app_record = {}
    app_path = directory / "app.md"
    if app_path.is_file():
        app_record, _body = read_markdown(app_path)
    pages = [record for _path, record in iter_records(directory, "page")]
    regions = [record for _path, record in iter_records(directory, "region")]
    controls = [record for _path, record in iter_records(directory, "control")]
    interactions = [record for _path, record in iter_records(directory, "interaction")]
    index = {
        "schema_version": INDEX_VERSION,
        "generated_at": utc_now(),
        "application": app_record,
        "pages": pages,
        "regions": regions,
        "controls": controls,
        "interactions": interactions,
    }
    index_path = directory / ".easy_uiauto" / "index.json"
    index_path.parent.mkdir(parents=True, exist_ok=True)
    index_path.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    return index


def load_index(directory: Path, rebuild: bool = False) -> dict:
    path = directory / ".easy_uiauto" / "index.json"
    if not rebuild and path.is_file():
        try:
            index = json.loads(path.read_text(encoding="utf-8"))
            if index.get("schema_version") == INDEX_VERSION:
                return index
        except (OSError, json.JSONDecodeError):
            pass
    return rebuild_index(directory)


def list_interactions(directory: Path, command: str = "", limit: int = 50) -> list[dict]:
    records = load_index(directory).get("interactions", [])
    if command:
        records = [item for item in records if item.get("command") == command]
    records.sort(key=lambda item: item.get("updated_at", ""), reverse=True)
    return records[: max(1, min(int(limit), 500))]


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
                    control.get("status") == "quarantined" for control in index.get("controls", [])
                ),
                "semantic_verified": sum(
                    control.get("semantic_status") in {"verified", "manual"}
                    for control in index.get("controls", [])
                ),
                "semantic_uncertain": sum(
                    control.get("semantic_status") == "uncertain"
                    for control in index.get("controls", [])
                ),
                "interactions": len(index.get("interactions", [])),
            }
        )
    return result


def _search_text(record: dict) -> str:
    values = [
        record.get("id", ""),
        record.get("semantic_name", ""),
        record.get("intent", ""),
        record.get("description", ""),
        record.get("semantic_role", ""),
        record.get("risk", ""),
        record.get("semantic_ambiguity", ""),
        record.get("name", ""),
        record.get("control_type", ""),
        record.get("automation_id", ""),
        record.get("page_id", ""),
        record.get("region_id", ""),
        record.get("notes", ""),
        " ".join(record.get("tags", [])),
        " ".join(record.get("actions", [])),
        " ".join(record.get("aliases", [])),
        " ".join(record.get("semantic_evidence", [])),
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
        if control.get("semantic_status") not in {"verified", "manual"}:
            continue
        if page_id and control.get("page_id") != page_id:
            continue
        for action in control.get("actions", []):
            commands.append(
                {
                    "command": f"{control['command']}.{action}",
                    "control_id": control["id"],
                    "semantic_name": control.get("semantic_name") or control.get("name"),
                    "intent": control.get("intent"),
                    "description": control.get("description", ""),
                    "aliases": control.get("aliases", []),
                    "semantic_confidence": control.get("semantic_confidence", 0),
                    "semantic_source": control.get("semantic_source", "unknown"),
                    "function_verification": control.get("function_verification", {}),
                    "risk": control.get("risk", "unknown"),
                    "requires_confirmation": control.get("requires_confirmation", False),
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
