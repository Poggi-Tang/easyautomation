"""Build a management tree without changing locators or navigation actions."""

from __future__ import annotations

from collections import defaultdict

from .protocol import normalize_location


def region_key(record: dict) -> str:
    page = record.get("page_id", "unknown")
    region = str(record.get("region_id") or "unassigned")
    return region if region.startswith(page+".") else page+"."+region


def _path(record):
    try:
        steps = normalize_location(record["location"])["Xpath"]
        return tuple(tuple(str(step.get(key) or "") for key in
                           ("ControlType", "Name", "AutomationId", "ClassName", "foundIndex"))
                     for step in steps)
    except (KeyError, TypeError, ValueError, AttributeError):
        return ()


def _relations(controls):
    records = {r["id"]: r for r in controls}
    paths = {identifier: _path(r) for identifier, r in records.items()}
    candidates = defaultdict(list)
    for identifier, record in records.items():
        if paths[identifier]:
            candidates[(record.get("page_id"), region_key(record), paths[identifier])].append(identifier)
    parents, sources = {}, {}
    for identifier, record in records.items():
        manual = record.get("hierarchy_parent", "auto")
        if not isinstance(manual, str):
            parents[identifier], sources[identifier] = None, "invalid-parent"
            continue
        if manual != "auto":
            parents[identifier] = manual if manual in records else None
            sources[identifier] = "manual" if not manual or manual in records else "missing-parent"
            continue
        navigation = record.get("navigation_path", [])
        parent = navigation[-1] if isinstance(navigation, list) and navigation else None
        if isinstance(parent, str) and parent in records and parent != identifier:
            parents[identifier], sources[identifier] = parent, "navigation"
            continue
        parents[identifier], sources[identifier] = None, "region"
        path = paths[identifier]
        for depth in range(len(path)-1, 0, -1):
            matches = candidates.get((record.get("page_id"), region_key(record), path[:depth]), [])
            if len(matches) == 1:
                parents[identifier], sources[identifier] = matches[0], "uia"
                break
    return records, parents, sources


def build(controls: list[dict]) -> dict:
    records, parents, sources = _relations(controls)
    warnings, finished = [], set()
    for identifier in records:
        trail, positions = [], {}
        current = identifier
        while current is not None and current not in finished:
            if current in positions:
                for cyclic in trail[positions[current]:]:
                    parents[cyclic], sources[cyclic] = None, "invalid-cycle"
                    warnings.append({"control_id": cyclic, "reason": "hierarchy-cycle"})
                break
            positions[current] = len(trail)
            trail.append(current)
            current = parents[current]
        finished.update(trail)
    regions = {}
    for identifier in records:
        current, trail = identifier, []
        while current not in regions:
            trail.append(current)
            parent = parents[current]
            if parent is None:
                regions[current] = region_key(records[current])
                break
            current = parent
        inherited = regions[current]
        for child in reversed(trail):
            regions[child] = inherited
    children = defaultdict(list)
    for identifier, parent in parents.items():
        if parent is not None:
            children[parent].append(identifier)
    return {"parents": parents, "sources": sources, "regions": regions,
            "children": dict(children), "warnings": warnings}


def validate_parent(controls: list[dict], control_id: str, parent: str) -> None:
    if not isinstance(parent, str):
        raise ValueError("上级控件必须为控件 ID、auto 或空字符串")
    records = {r["id"]: r for r in controls}
    if control_id not in records:
        raise ValueError("控件记录不存在")
    if parent not in {"auto", ""} and parent not in records:
        raise ValueError("所选上级控件不存在或不属于当前软件")
    proposed = [{**r, "hierarchy_parent": parent} if r["id"] == control_id else r for r in controls]
    _records, parents, _sources = _relations(proposed)
    current, seen = control_id, set()
    while current is not None:
        if current in seen:
            raise ValueError("控件层级不能形成循环，不能把自身或下级设为上级")
        seen.add(current)
        current = parents[current]
