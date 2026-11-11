"""Regions tied to a learned expanded state, never to a click target."""

from __future__ import annotations

from . import control_hierarchy, knowledge


def is_dynamic(record: dict) -> bool:
    return record.get("source") in {"dynamic-container", "manual-dynamic", "local-visual-container"}


def _same_state(first, second):
    return all(first.get(key, default) == second.get(key, default) for key, default in (
        ("app_id", ""), ("page_id", ""), ("navigation_path", []), ("window_scope", {"kind": "main"}))) and (
        first.get("visual_observation", {}).get("context_image") == second.get("visual_observation", {}).get("context_image"))


def _region_sources(controls, identifiers):
    """Collapse an automatic wrapper only with structural AND state evidence."""
    records = {r["id"]: r for r in controls}
    paths = {i: control_hierarchy._path(r) for i, r in records.items()}
    commands = {i: paths[i] for i, r in records.items() if r.get("control_type") in {
        "MenuItemControl", "ButtonControl", "CheckBoxControl", "RadioButtonControl", "HyperlinkControl"}}
    contents = {i: {path for child, path in commands.items() if len(path) > len(paths[i])
                    and path[:len(paths[i])] == paths[i] and _same_state(records[i], records[child])}
                for i in identifiers}
    result = {i: i for i in identifiers}
    for identifier in sorted(identifiers, key=lambda i: (len(paths[i]), i)):
        record, path = records[identifier], paths[identifier]
        if record.get("region_behavior") == "region" or not path: continue
        n = record.get("visual_observation", {}).get("normalized_rect", {})
        same_path = [other for other in identifiers if other < identifier and paths[other] == path
                     and _same_state(record, records[other])
                     and n and all(abs(n[k]-records[other].get("visual_observation", {}).get("normalized_rect", {}).get(k, -10)) < .001
                                   for k in ("left", "top", "right", "bottom"))]
        if same_path:
            result[identifier] = result[min(same_path)]
            continue
        ancestors = []
        for other in identifiers:
            if other == identifier or not paths[other] or len(paths[other]) >= len(path): continue
            if path[:len(paths[other])] != paths[other] or not _same_state(record, records[other]): continue
            outer = records[other].get("visual_observation", {}).get("normalized_rect", {})
            if not n or not outer: continue
            contained = all(n[k] >= outer[k]-.001 for k in ("left", "top")) and all(
                n[k] <= outer[k]+.001 for k in ("right", "bottom"))
            if contained and (len(contents[identifier]) <= 1 or contents[identifier] == contents[other]):
                ancestors.append(other)
        if ancestors:
            parent = max(ancestors, key=lambda i: len(paths[i]))
            result[identifier] = result[parent]
    return result


def synchronize(directory, controls: list[dict], observed_ids=None) -> int:
    """Materialize known popup containers; only learning updates their pixels."""
    existing = {r.get("source_control_id"): r for _p, r in knowledge.iter_records(directory, "region")
                if r.get("source") == "dynamic-container" and isinstance(r.get("source_control_id"), str)}
    records = {r["id"]: r for r in controls}
    eligible = {i for i, r in records.items() if r.get("navigation_path") and r.get("visual_observation", {}).get("normalized_rect")
                and (r.get("control_type") == "MenuControl" or r.get("region_behavior") == "region")
                and r.get("region_behavior") != "control"}
    sources = _region_sources(controls, eligible)
    changed = 0
    for control in controls:
        identifier = control["id"]
        behavior = control.get("region_behavior", "auto")
        navigation = control.get("navigation_path")
        if (behavior == "control" or not isinstance(navigation, list) or not navigation
                or any(not isinstance(i, str) or i not in records or i == identifier for i in navigation)):
            continue
        if behavior != "region" and control.get("control_type") != "MenuControl":
            continue
        if sources.get(identifier, identifier) != identifier:
            # Keep the raw control and manual regions; retire only redundant auto metadata.
            if identifier in existing:
                from uuid import uuid4
                from . import image_storage
                old = existing[identifier]
                retained = existing.get(sources[identifier], {}).get("id") or knowledge.stable_id("dynamic-region", sources[identifier])
                for _child_path, child in knowledge.iter_records(directory, "region"):
                    if child.get("parent_region_id") == old["id"]:
                        knowledge.save_region(directory, {**child, "parent_region_id": retained})
                path = directory / "regions" / (knowledge.slugify(old["id"])+".md")
                target = directory / ".easy_uiauto" / "archive" / uuid4().hex / "regions" / path.name
                if image_storage._safe_path(directory, target) is None:
                    raise ValueError("冗余区域归档路径无效")
                target.parent.mkdir(parents=True, exist_ok=True)
                path.replace(target)
                changed += 1
            continue
        observation = control.get("visual_observation", {})
        if not observation.get("normalized_rect") or not observation.get("context_image"):
            continue
        previous = existing.get(identifier, {})
        if previous and observed_ids is not None and identifier not in observed_ids:
            continue
        label = control.get("semantic_name") or control.get("name") or "展开"
        default_name = f"{label} 子菜单" if control.get("control_type") == "MenuControl" else f"{label} 区域"
        record = {**previous, "id": previous.get("id") or knowledge.stable_id("dynamic-region", identifier),
                  "app_id": control.get("app_id"), "page_id": control.get("page_id"),
                  "name": previous.get("name") or default_name,
                  "source": "dynamic-container", "geometry_role": "visual-hint-only",
                  "coordinate_space": "observation", "source_control_id": identifier,
                  "parent_control_id": navigation[-1],
                  "navigation_path": list(navigation),
                  "window_scope": control.get("window_scope", {"kind": "main"}),
                  "normalized_rect": observation["normalized_rect"],
                  "visual_observation": observation, "image": control.get("image", "")}
        record.pop("kind", None)
        record.pop("updated_at", None)
        comparable = {k: v for k, v in previous.items() if k not in {"kind", "updated_at"}}
        if record != comparable:
            knowledge.save_region(directory, record)
            changed += 1
    return changed


def build_tree(controls: list[dict], regions: list[dict]) -> dict:
    """Use typed node IDs so a popup region replaces its source display node."""
    base = control_hierarchy.build(controls)
    records = {r["id"]: r for r in controls}
    dynamic = {r["id"]: r for r in regions if is_dynamic(r)}
    containers = {r["source_control_id"]: r["id"] for r in dynamic.values()
                  if isinstance(r.get("source_control_id"), str) and r.get("source_control_id") in records
                  and records[r["source_control_id"]].get("region_behavior") != "control"}
    sources = _region_sources(controls, set(containers))
    containers = {i: containers[primary] for i, primary in sources.items()}
    # Even after migration, transparent native wrappers map to the retained region.
    eligible = set(containers) | {i for i, r in records.items() if r.get("control_type") == "MenuControl"
        and r.get("navigation_path") and r.get("region_behavior") != "control"}
    for identifier, primary in _region_sources(controls, eligible).items():
        if primary in containers:
            containers[identifier] = containers[primary]
    dynamic = {i: r for i, r in dynamic.items() if r.get("source") != "dynamic-container"
               or isinstance(r.get("source_control_id"), str) and containers.get(r["source_control_id"]) == i}
    parents = {}

    def control_node(identifier):
        return "region:"+containers[identifier] if identifier in containers else "control:"+identifier

    for identifier in records:
        if identifier in containers:
            continue
        parent = base["parents"][identifier]
        parents[control_node(identifier)] = control_node(parent) if parent else "region:"+base["regions"][identifier]
        record = records[identifier]
        if record.get("hierarchy_parent", "auto") != "auto":
            continue
        if parent in containers:
            source = records[parent]
            if (record.get("navigation_path") != source.get("navigation_path")
                    or record.get("window_scope", {"kind": "main"}) != source.get("window_scope", {"kind": "main"})):
                navigation = record.get("navigation_path", [])
                if not isinstance(navigation, list): navigation = []
                opener = navigation[-1] if navigation else None
                parents[control_node(identifier)] = (control_node(opener) if isinstance(opener, str) and opener in records else
                                                     "region:"+control_hierarchy.region_key(record))
        # A native ancestor from this expanded state is stronger than its opener.
        path = control_hierarchy._path(record)
        matches = []
        for source_id, region_id in containers.items():
            source = records[source_id]
            prefix = control_hierarchy._path(source)
            if (prefix and len(path) > len(prefix) and path[:len(prefix)] == prefix
                    and record.get("page_id") == source.get("page_id")
                    and record.get("navigation_path") == source.get("navigation_path")
                    and record.get("window_scope", {"kind": "main"}) == source.get("window_scope", {"kind": "main"})):
                matches.append((len(prefix), region_id))
        if matches:
            deepest = max(length for length, _r in matches)
            unique = set(i for length, i in matches if length == deepest)
            if len(unique) == 1:
                parents[control_node(identifier)] = "region:"+next(iter(unique))
        local_matches=[]
        for region_id,region in dynamic.items():
            prefix=control_hierarchy._path(region)
            if (region.get("source")=="local-visual-container" and prefix and len(prefix)<len(path)
                    and path[:len(prefix)]==prefix and region.get("page_id")==record.get("page_id")
                    and region.get("navigation_path")==record.get("navigation_path")
                    and region.get("window_scope",{"kind":"main"})==record.get("window_scope",{"kind":"main"})):
                local_matches.append((len(prefix),region_id))
        if local_matches:
            deepest=max(length for length,_r in local_matches)
            unique={r for length,r in local_matches if length==deepest}
            if len(unique)==1:parents[control_node(identifier)]="region:"+next(iter(unique))
    for identifier, region in dynamic.items():
        if region.get("source") == "dynamic-container":
            source_id = region["source_control_id"]
            parent = base["parents"].get(source_id)
            if records[source_id].get("hierarchy_parent", "auto") == "auto":
                path = control_hierarchy._path(records[source_id])
                ancestors = [other for other in containers if containers[other] != identifier
                    and control_hierarchy._path(records[other])
                    and len(control_hierarchy._path(records[other])) < len(path)
                    and path[:len(control_hierarchy._path(records[other]))] == control_hierarchy._path(records[other])
                    and _same_state(records[source_id], records[other])]
                if ancestors:
                    parent = max(ancestors, key=lambda i: len(control_hierarchy._path(records[i])))
            parents["region:"+identifier] = control_node(parent) if parent else "region:"+base["regions"][source_id]
        else:
            parent = region.get("parent_region_id")
            opener = region.get("parent_control_id")
            parents["region:"+identifier] = ("region:"+parent if isinstance(parent, str) and parent in dynamic else
                                              control_node(opener) if isinstance(opener, str) and opener in records else None)

    # Manually drawn regions partition only this container's existing controls.
    for identifier in records:
        node = "control:"+identifier
        parent = parents.get(node, "") or ""
        if not parent.startswith("region:") or parent[7:] not in dynamic:
            continue
        region = dynamic[parent[7:]]
        source_id = region.get("source_control_id")
        source = records.get(source_id, {})
        n = records[identifier].get("visual_observation", {}).get("normalized_rect")
        frame = records[identifier].get("visual_observation", {}).get("context_image")
        if not n or frame != source.get("visual_observation", {}).get("context_image"):
            continue
        current = parent[7:]
        seen = set()
        while current not in seen:
            seen.add(current)
            children = [r for r in dynamic.values() if r.get("parent_region_id") == current
                        and r.get("visual_observation", {}).get("context_image") == frame]
            if not children:
                break
            from .manual_layout import assign_region
            selected, _reason = assign_region(n, children)
            if selected == "unassigned":
                break
            parents[node] = "region:"+selected
            current = selected

    # Corrupt metadata must not create a recursive GUI or discard learned controls.
    roots = {"region:"+r["id"] for r in regions if not is_dynamic(r)}
    for node in list(parents):
        trail, current = set(), node
        while current in parents and current not in trail:
            trail.add(current)
            current = parents[current]
        if current in trail:
            parents[node] = None
    children = {}
    for node, parent in parents.items():
        children.setdefault(parent, []).append(node)
    return {"parents": parents, "children": children, "containers": containers,
            "dynamic": dynamic, "static_roots": roots, "base": base}
