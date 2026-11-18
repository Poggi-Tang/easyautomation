"""Archive unused legacy labels only when a fresh scan proves their ownership."""
from __future__ import annotations

import os

from . import knowledge, local_learning
from .diagram_geometry import is_text
from .protocol import normalize_location


def _path(location):
    try:
        steps = normalize_location(location)["Xpath"]
    except (TypeError, ValueError):
        return []
    return [{"match_mode": "identity", **step} for step in steps]


def _matches(record, component):
    if not component.get("native_evidence"):
        return False
    if record.get("control_type") in {"TextControl", "ImageControl"}:
        return (component.get("component_binding") == "passive-native-child"
                and bool(_path(record.get("location", {})))
                and _path(record["location"]) == _path(component.get("native_location", {})))
    if record.get("control_type") != "VisualElement":
        return False
    reason = record.get("verification", {}).get("reason")
    if reason != "shared_native_control" and not (
            reason == "geometry_rejected" and component.get("component_binding") == "same-native-widget"):
        return False
    # Old visual failures have no native identity. Require the same observation
    # and practically identical geometry, never proximity or a shared name alone.
    from .scanner import _rect_iou
    before, after = record.get("normalized_rect", {}), component.get("normalized_rect", {})
    if not all(key in before and key in after for key in ("left", "top", "right", "bottom")):
        return False
    same_kind = (is_text(record) and is_text(component)) or (
        record.get("entity_type") == component.get("entity_type"))
    return (same_kind and record.get("observed_value", "") == component.get("observed_value", "")
            and _rect_iou(before, after) >= .98)


def _protection(record):
    if (record.get("semantic_source") == "manual"
            or record.get("description_source") in {"manual", "interaction-observation"}
            or record.get("region_source") == "manual-selection"
            or record.get("hierarchy_parent", "auto") != "auto"
            or record.get("entry_steps") or record.get("navigation_path") or record.get("region_behavior")):
        return "manual-or-navigation-data"
    function = record.get("function_verification", {})
    if (knowledge.runtime_quarantined(record) or record.get("last_click_execution")
            or function.get("executed") or function.get("effect_record")
            or function.get("status") == "human-confirmed"):
        return "execution-history"
    return ""


def reconcile(directory, previous, observed, scan_id):
    """Return archived mappings and protected duplicates for the scan report."""
    report = {"archived": [], "retained": []}
    current = {r["id"]: (p, r) for p, r in knowledge.iter_records(directory, "control")}
    owners = [r for _p, r in current.values()
              if r.get("scan_id") == scan_id and r.get("scan_strategy") == "local-visual"
              and r.get("control_type") in local_learning._ATOMIC_CONTROLS - {"TextControl"}
              and r.get("semantic_source") != "manual" and not knowledge.runtime_quarantined(r)]
    candidates = {}
    for old in previous:
        identifier = old["id"]
        if (identifier in observed or identifier not in current
                or old.get("scan_strategy") != "local-visual"
                or old.get("control_type") not in {"TextControl", "ImageControl", "VisualElement"}):
            continue
        matches = [owner for owner in owners
                   if all(old.get(key) == owner.get(key) for key in ("app_id", "page_id", "region_id"))
                   and old.get("window_scope", {"kind": "main"}) == owner.get("window_scope", {"kind": "main"})
                   and old.get("navigation_path", []) == owner.get("navigation_path", [])
                   and any(_matches(old, part) for part in owner.get("visual_components", []))]
        if len(matches) != 1:
            continue
        mapping = {"control_id": identifier, "owner_control_id": matches[0]["id"]}
        reason = _protection(old)
        if reason:
            report["retained"].append({**mapping, "reason": reason})
        else:
            candidates[identifier] = mapping
    if not candidates:
        return report

    # Read references once, including durable deep-learning JSON and user guides.
    # Generated catalogs/indexes are rebuilt after reconciliation.
    generated = {"index.md", "commands.md", "operations/UI-CLI.md"}
    blocked = {}
    for folder, subdirs, names in os.walk(directory):
        subdirs[:] = [name for name in subdirs if name not in {".easy_uiauto", ".git", "images", "analysis"}]
        for name in names:
            path = directory / os.path.relpath(folder, directory) / name
            if (path.suffix.lower() not in {".md", ".json", ".jsonl", ".yaml", ".yml"}
                    or path.relative_to(directory).as_posix() in generated):
                continue
            try:
                text = path.read_text(encoding="utf-8")
            except (OSError, UnicodeError):
                # If dependency inspection is incomplete, keep the originals.
                blocked.update({identifier: "reference-check-failed" for identifier in candidates})
                continue
            for identifier in candidates:
                own_path, record = current[identifier]
                if path == own_path:
                    continue
                if identifier in text or (record.get("command") and record["command"] in text):
                    blocked[identifier] = "referenced"
    for identifier, mapping in candidates.items():
        if identifier in blocked:
            report["retained"].append({**mapping, "reason": blocked[identifier]})
            continue
        local_learning.archive_record(directory, current[identifier][0], scan_id,
            reason={**mapping, "reason": "native-component-evidence", "scan_id": scan_id})
        report["archived"].append(mapping)
    return report
