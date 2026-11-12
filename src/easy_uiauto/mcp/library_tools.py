"""Agent access to the GUI's existing vault, learning workers and workflow format."""
from __future__ import annotations

import atexit
from copy import deepcopy
import threading
import time

from ..knowledge_library import KnowledgeLibrary
from ..library_tasks import LibraryTask
from .. import workflows
from . import dynamic_regions, image_storage, knowledge, native_windows


SECTIONS = {"pages", "regions", "controls", "references", "usage", "workflows"}
OPERATIONS = {"learn-region", "expand-control", "run-workflow", "supplement-descriptions"}


def get_library(app_id="", *, vault="", section="overview", record_id="", page_id="",
                region_id="", offset=0, limit=50, text_offset=0, text_limit=12000):
    """Paginate records and document bodies without dropping manual annotations."""
    if offset < 0 or not 1 <= limit <= 200 or text_offset < 0 or not 1 <= text_limit <= 30000:
        raise ValueError("offset 必须非负，limit 为 1..200，text_limit 为 1..30000")
    library = KnowledgeLibrary(vault or None)
    result = dict(ok=True, vault=str(library.root), app_id=app_id, section=section)
    if not app_id:
        records = library.apps()
    else:
        snapshot = library.read(app_id)
        if section == "overview":
            return {**result, "application": snapshot["application"],
                    "counts": {key: len(snapshot[key]) for key in SECTIONS},
                    "sections": sorted(SECTIONS)}
        if section not in SECTIONS:raise ValueError("未知知识分类："+section)
        records = snapshot[section]
        if page_id:records = [r for r in records if r.get("page_id", r.get("id")) == page_id]
        if region_id:
            region = next((r for r in snapshot["regions"] if r["id"] == region_id), None)
            if region is None:raise ValueError("区域 ID 不存在，请读取 regions 分类")
            local = region_id.removeprefix(region["page_id"]+".")
            records = [r for r in records if r.get("page_id") == region["page_id"]
                       and r.get("region_id", r.get("id")) in {region_id, local}]
        if record_id:
            records = [r for r in records if r.get("id") == record_id or r.get("document_path") == record_id]
            if not records:raise ValueError("指定记录不存在或不在所选范围内")
    records = sorted(records, key=lambda r: str(r.get("id") or r.get("document_path", "")))
    page = deepcopy(records[offset:offset+limit])
    for record in page:
        if "body" in record:
            body = record.pop("body")
            record["body_length"] = len(body)
            if record_id:
                record["body"] = body[text_offset:text_offset+text_limit]
                record["text_offset"] = text_offset
                record["next_text_offset"] = text_offset+text_limit if text_offset+text_limit < len(body) else None
        if section == "controls" and app_id:
            record["effective_entry_steps"] = library.entry_steps(app_id, record, snapshot=snapshot)
            record["effective_region_state"] = library.record_state(app_id, record, snapshot=snapshot)
    return {**result, "records": page, "total": len(records), "offset": offset,
            "next_offset": offset+limit if offset+limit < len(records) else None}


def readiness():
    from ..image_models import inspect_models
    from .configuration import load_gui_configuration
    report = inspect_models()
    settings = load_gui_configuration()
    return {**report, "ready": report["ok"], "engine": "local-visual",
            "reference_api_configured": all(settings.get(k) for k in ("api_url", "api_key", "api_model")),
            "setup_instruction": "" if report["ok"] else "打开 GUI 的标注模型面板，配置模型路径或下载 / 修复模型。"}


def _window(application, handle):
    if isinstance(handle, bool) or not isinstance(handle, int) or handle < 0:
        raise ValueError("window_handle 必须是非负整数")
    matches = [w for w in native_windows.matching_windows(application.get("window_name", ""), handle=handle)
               if w["visible"] and w["class_name"] == application.get("window_class")]
    if len(matches) != 1:
        raise ValueError("无法唯一确定已绑定的软件窗口；请从 list_windows 选择主窗口并传入 window_handle")
    return matches[0]


def _request(operation, app_id, *, vault="", region_id="", control_id="", workflow_id="",
             window_handle=0, parameters=None, max_controls=3000):
    if operation not in OPERATIONS:raise ValueError("未知知识库任务")
    if isinstance(max_controls, bool) or not isinstance(max_controls, int) or not 1 <= max_controls <= 20000:
        raise ValueError("max_controls 必须为 1..20000")
    library = KnowledgeLibrary(vault or None)
    snapshot = library.read(app_id)
    payload = dict(vault=str(library.root), app_id=app_id, capture_handshake=False)
    if operation == "supplement-descriptions":
        ids = None
        if control_id:
            if not any(r["id"] == control_id for r in snapshot["controls"]):raise ValueError("控件不存在")
            ids = [control_id]
        elif region_id:
            region = next((r for r in snapshot["regions"] if r["id"] == region_id), None)
            if region is None:raise ValueError("区域不存在")
            local = region_id.removeprefix(region["page_id"]+".")
            ids = [r["id"] for r in snapshot["controls"] if r.get("page_id") == region["page_id"]
                   and r.get("region_id") in {region_id, local}]
        return {**payload, "operation": operation, "control_ids": ids}
    window = _window(snapshot["application"], window_handle)
    payload.update(window=window, window_name=window["title"])
    if operation == "run-workflow":
        flow = next((r for r in snapshot["workflows"] if r["id"] == workflow_id), None)
        if flow is None:raise ValueError("组合动作不存在，请先保存")
        report = workflows.readiness(flow, snapshot["controls"])
        if not report["ready"]:raise ValueError("组合动作尚未完善："+"；".join(i["message"] for i in report["issues"]))
        if parameters is not None and not isinstance(parameters, dict):raise ValueError("parameters 必须为对象")
        workflows.substitute(flow["steps"], {**flow.get("parameters", {}), **(parameters or {})})
        return {**payload, "operation": "workflow", "flow_id": workflow_id, "parameters": parameters or {}}
    if operation == "learn-region":
        region = next((r for r in snapshot["regions"] if r["id"] == region_id), None)
        if region is None:raise ValueError("请选择已保存的区域 ID")
        if dynamic_regions.is_dynamic(region):
            raise ValueError("动态区域请使用 expand-control，传入其 source_control_id 或展开入口 control_id")
        if region.get("source") != "manual" or region.get("classification_only"):
            raise ValueError("只学习 GUI 中已划分的人工区域")
        scope = library.scoped(region["page_id"])
        layout = scope.learning_layout(app_id, window, region_id=region_id)
        payload.update(operation="learn", region_id=region_id, layout=layout)
    else:
        control = next((r for r in snapshot["controls"] if r["id"] == control_id), None)
        if control is None:raise ValueError("请选择已入库的展开入口控件")
        library.click_plan(app_id, control_id)
        layout = library.scoped(control["page_id"]).learning_layout(app_id, window)
        payload.update(operation="expand", control_id=control_id, layout=layout)
    return {**payload, "semantic_mode": "native", "max_controls": max_controls}


class LibraryJobs:
    """Own one worker at a time, including after the MCP request has returned."""
    def __init__(self):
        self.lock = threading.RLock()
        self.jobs = {}
        self.latest = ""
        self.active = ""
        self.thread = None

    def require_idle(self):
        if self.active:
            raise ValueError(f"知识库任务仍在运行：{self.active}；请查询 get_ui_library_task，不要重复启动")

    def start(self, operation, app_id, **kwargs):
        with self.lock:
            self.require_idle()
            payload = _request(operation, app_id, **kwargs)
            worker = LibraryTask(payload, timeout=3600 if operation == "supplement-descriptions" else 600)
            identifier = "library-"+worker.folder.name
            entry = dict(task_id=identifier, operation=operation, app_id=app_id, vault=payload["vault"],
                         state="running", worker_running=True, elapsed_seconds=0, message="任务已启动",
                         worker=worker, cancel_requested=False, started=time.monotonic())
            self.jobs[identifier] = entry
            self.latest = self.active = identifier
            # Completion status remains available without retaining temporary screenshots.
            while len(self.jobs) > 20:self.jobs.pop(next(iter(self.jobs)))
            self.thread = threading.Thread(target=self._watch, args=(identifier,), daemon=True,
                                           name="easy-uiauto-library")
            self.thread.start()
            return self._public(entry)

    @staticmethod
    def _public(entry):
        result = deepcopy({key: value for key, value in entry.items() if key not in {"worker", "started"}})
        outcome = result.get("result") or {}
        result.update(ok=result["state"] not in {"failed", "cancelled"} and outcome.get("ok") is not False,
                      status_tool="get_ui_library_task", stop_tool="stop_ui_library_task")
        return result

    def _watch(self, identifier):
        while True:
            with self.lock:
                entry = self.jobs[identifier]
                worker = entry.get("worker")
                if worker is None:
                    return
                try:
                    status = worker.poll()
                except Exception as error:
                    status = dict(state="failed", error=str(error), worker_running=False)
                entry.update(status)
                if status["worker_running"]:
                    # A process can publish completed before it releases locks and exits.
                    entry["state"] = "cancelling" if entry["cancel_requested"] else "running"
                else:
                    try:worker.close()
                    except Exception as error:entry["cleanup_error"] = str(error)
                    entry.pop("worker")
                    self.active = ""
                    return
            time.sleep(.2)

    def status(self, task_id="", *, stop=False):
        with self.lock:
            identifier = task_id or self.latest
            if identifier not in self.jobs:raise ValueError("知识库任务不存在；MCP 重启后请先检查已归档记录")
            entry = self.jobs[identifier]
            if stop and entry.get("worker"):
                entry["worker"].cancel();entry["cancel_requested"] = True;entry["state"] = "cancelling"
            result = self._public(entry)
            if entry.get("worker"):result["elapsed_seconds"] = round(time.monotonic()-entry["started"], 1)
            return result

    def close(self):
        with self.lock:
            for entry in self.jobs.values():
                worker = entry.pop("worker", None)
                if worker:
                    try:
                        worker.cancel()
                        worker.close()
                    except Exception:pass
                    entry.update(state="cancelled", worker_running=False, cancel_requested=True)
            self.active = ""


jobs = LibraryJobs()
atexit.register(jobs.close)


def save_workflow(app_id, workflow, *, vault=""):
    """Upsert a definition; only execution workers may produce run evidence."""
    if not isinstance(workflow, dict):raise ValueError("workflow 必须为对象")
    allowed = {"id", "keyword", "name", "precondition", "parameters", "steps", "source_case", "source_instructions"}
    unknown = set(workflow)-allowed
    if unknown:raise ValueError("仅接受流程定义，状态与运行结果由程序生成；不支持字段："+", ".join(sorted(unknown)))
    library = KnowledgeLibrary(vault or None)
    directory = library.directory(app_id)
    with jobs.lock:
        jobs.require_idle()
        with image_storage._vault_lock(library.root):
            snapshot = library.read(app_id)
            keyword = workflow.get("keyword", "")
            if not isinstance(keyword, str) or not keyword.strip():raise ValueError("流程关键字不能为空")
            keyword = keyword.strip()
            identifier = workflow.get("id")
            existing = next((r for r in snapshot["workflows"] if r["id"] == identifier), None) if identifier else None
            if identifier and existing is None:raise ValueError("指定流程 ID 不存在；新建时省略 id")
            named = next((r for r in snapshot["workflows"] if r["keyword"] == keyword), None)
            if existing and named and existing["id"] != named["id"]:raise ValueError("已有其他同名组合动作")
            existing = existing or named
            definition = {**(existing or {}), **deepcopy(workflow), "keyword": keyword,
                          "id": existing["id"] if existing else knowledge.stable_id("workflow", keyword),
                          "kind": "workflow", "app_id": app_id}
            definition.setdefault("parameters", {});definition.setdefault("precondition", "软件已启动")
            workflows.validate(definition)
            controls = {r["id"]: r for r in snapshot["controls"]}
            for _, target in workflows.targets(definition):
                target.setdefault("label", "")
                bound = target.get("control_id")
                if bound and bound not in controls:raise ValueError("控件 ID 不属于当前软件或已删除："+bound)
            definition = workflows.suggest_bindings(definition, snapshot["controls"])
            saved = library.save_workflow(app_id, definition)
    return dict(ok=True, action="updated" if existing else "created", vault=str(library.root),
                document_path=str(directory/"workflows"/(saved["id"]+".md")), workflow=saved,
                verified=False, message="已保存组合动作；ready 表示定义可运行，实际验证请查看 last_run。")
