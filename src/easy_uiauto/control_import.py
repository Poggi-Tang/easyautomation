"""Import legacy control observations into the shared Markdown knowledge library."""
from __future__ import annotations

from copy import deepcopy
import json
from pathlib import Path

from .mcp import image_storage, knowledge
from .mcp.protocol import normalize_location
from .mcp.control_hierarchy import _path


def prepare(library, source, *, app_id="", page_id="", region_id=""):
    """Resolve an observation without inventing application or state membership."""
    if not isinstance(source, dict):
        raise ValueError("控件记录必须为对象")
    for key in ("app_id", "page_id", "region_id", "control_id", "alias", "app_name", "window_name", "semantic_name", "description", "notes"):
        if key in source and not isinstance(source[key], str):raise ValueError(key + " 必须为字符串")
    location = normalize_location(source.get("location") or source.get("LOCATION") or source)
    if not location.get("Xpath"):
        raise ValueError("控件记录必须包含完整 LOCATION / XPath")
    window = next((step for step in location["Xpath"] if step.get("ControlType") == "WindowControl"), {})
    title = source.get("window_name") or window.get("Name") or location.get("WindowName", "")
    window_class = window.get("ClassName", "")
    app_id = app_id or source.get("app_id", "")
    if not app_id:
        apps = [app for app in library.apps()
                if title and app.get("window_name") == title
                and (not window_class or not app.get("window_class") or app["window_class"] == window_class)]
        if len(apps) > 1:
            raise ValueError("窗口对应多个软件，请指定 app_id")
        if apps:
            app_id = apps[0]["id"]
        else:
            name = source.get("app_name") or title
            if not name:
                raise ValueError("缺少软件信息，请指定 app_id")
            app_id = knowledge.slugify(name)
            if any(app["id"] == app_id for app in library.apps()):
                raise ValueError("同名软件的窗口信息不同，请明确指定 app_id")
    directory = library.directory(app_id)
    exists = (directory / "app.md").is_file()
    snapshot = library.read(app_id) if exists else dict(controls=[], regions=[], pages=[])
    page_id = page_id or source.get("page_id", "")
    region_id = region_id or source.get("region_id", "")
    region = next((r for r in snapshot["regions"] if r["id"] == region_id), None) if region_id else None
    if region_id and region_id != "unassigned" and region is None:
        raise ValueError("所属区域不存在，请使用 get_ui_library 返回的完整区域 ID")
    if region:
        if page_id and page_id != region["page_id"]:
            raise ValueError("区域与页面状态不一致")
        page_id = region["page_id"]
    identifier = source.get("control_id", "")
    path = _path({"location":location})
    candidates = [r for r in snapshot["controls"]
                  if (r["id"] == identifier if identifier else _path(r) == path)
                  and (not page_id or r.get("page_id") == page_id)
                  and (not region_id or r.get("region_id") == region_id)
                  and (not source.get("window_scope") or r.get("window_scope", {"kind":"main"}) == source["window_scope"])]
    if identifier and not candidates:
        raise ValueError("指定控件不存在或不属于指定区域 / 页面")
    if len(candidates) > 1:
        raise ValueError("定位对应多个状态或入口，请指定 page_id、region_id 或 control_id")
    old = candidates[0] if candidates else {}
    page_id = page_id or old.get("page_id", "")
    if not page_id:
        pages = snapshot["pages"]
        if len(pages) > 1:
            raise ValueError("软件有多个页面状态，请指定 page_id 或 region_id")
        page_id = pages[0]["id"] if pages else library.page_id
    library.scoped(page_id)  # Validate IDs before creating any files.
    if snapshot["pages"] and not any(p["id"] == page_id for p in snapshot["pages"]):
        raise ValueError("页面状态不存在，请先创建区域状态")
    page = next((p for p in snapshot["pages"] if p["id"] == page_id), {})
    region_scope = (region or {}).get("window_scope") or page.get("window_scope") or {"kind":"main"}
    scope = source.get("window_scope") or old.get("window_scope") or region_scope
    if not isinstance(scope, dict) or scope.get("kind") not in {"main", "owned"}:
        raise ValueError("window_scope 必须为 main 或 owned 窗口描述")
    if region and scope != region_scope:
        raise ValueError("控件窗口与所属区域不一致")
    route = (region or {}).get("navigation_path", old.get("navigation_path", []))
    identifier = old.get("id") or knowledge.stable_id("control", app_id, page_id, location, scope, route)
    record = deepcopy(old) if old else dict(
        id=identifier, kind="control", app_id=app_id, page_id=page_id,
        region_id=region_id or "unassigned", status="observed", location=location,
        name=location.get("Name", ""), class_name=location.get("ClassName", ""),
        control_type=location.get("ControlType", ""), automation_id=location.get("AutomationId", ""),
        semantic_name=source.get("semantic_name") or location.get("Name") or source.get("alias") or identifier,
        description=source.get("description") or source.get("notes", ""), semantic_source="import",
        semantic_status="unconfirmed", semantic_confidence=0, actions=[], supported_actions=[],
        verification={"location": "pending", "image": "pending"},
        function_verification={"status": "unknown", "executed": False},
        window_scope=scope, navigation_path=route,
    )
    if old and _path(old) != path:
        raise ValueError("导入不能改写已有控件定位；请使用控件属性维护接口")
    if region and not old:
        record["region_source"] = "manual-selection"
    alias = source.get("alias", "")
    record["aliases"] = list(dict.fromkeys([*record.get("aliases", []), *([alias] if alias else [])]))
    tags = source.get("tags", [])
    if isinstance(tags, str):tags = [v.strip() for v in tags.split(",") if v.strip()]
    if not isinstance(tags, list) or any(not isinstance(v, str) for v in tags):
        raise ValueError("tags 必须为字符串列表或逗号分隔字符串")
    record["tags"] = list(dict.fromkeys([*record.get("tags", []), *tags]))
    observation = {key: source[key] for key in ("alias", "source", "notes", "framework_id") if key in source}
    record["import_observation"] = observation
    app = dict(name=source.get("app_name") or title or app_id, window_name=title, window_class=window_class)
    return dict(record=record, application=app, app_id=app_id, exists=exists,
                action="updated" if old else "created")


def upsert(library, source, **scope):
    """Idempotently archive metadata; imported status can never bypass verification."""
    with image_storage._vault_lock(library.root):
        plan = prepare(library, source, **scope)
        directory = library.directory(plan["app_id"])
        if not plan["exists"]:
            app = plan["application"]
            library.bind(plan["app_id"], app["name"], dict(title=app["window_name"], class_name=app["window_class"]))
        knowledge.save_control(directory, plan["record"])
        library._refresh(directory)
        saved = knowledge.find_control_record(directory, plan["record"]["id"], include_quarantine=True)[1]
        return dict(ok=True, vault=str(library.root), app_id=plan["app_id"], id=saved["id"],
                    action=plan["action"], alias=source.get("alias", ""), record=saved,
                    message="已写入共享知识库；导入不会新增定位或业务验证证据。")


def import_records(library, source_path, *, dry_run=True, **scope):
    """Migrate a JSON/JSONL export without loading external Python modules or deleting originals."""
    path = Path(source_path).expanduser().resolve()
    text = path.read_text(encoding="utf-8-sig")
    rows = [json.loads(line) for line in text.splitlines() if line.strip()] if path.suffix.lower() == ".jsonl" else json.loads(text)
    if isinstance(rows, dict):rows = rows.get("records", [rows])
    if not isinstance(rows, list):raise ValueError("导入文件须为记录数组、records 数组或 JSONL")
    results = []
    with image_storage._vault_lock(library.root):
        for number, row in enumerate(rows, 1):
            try:
                if dry_run:
                    plan = prepare(library, row, **scope)
                    result = dict(ok=True, id=plan["record"]["id"], app_id=plan["app_id"], action=plan["action"])
                else:result = upsert(library, row, **scope)
            except (ValueError, TypeError, KeyError) as error:
                result = dict(ok=False, error=str(error))
            results.append({"row": number, **result})
    return dict(ok=all(r["ok"] for r in results), dry_run=dry_run, source=str(path), vault=str(library.root),
                total=len(rows), succeeded=sum(r["ok"] for r in results), results=results)
