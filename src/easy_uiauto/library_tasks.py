"""Isolated, cancellable jobs for the knowledge GUI; the UI never waits on UIA."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path
from uuid import uuid4


def _write(path: Path, value: dict):
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
    # Windows readers briefly hold the destination without FILE_SHARE_DELETE.
    # Retry only publishing the completed JSON, never the job's UI actions.
    for attempt in range(6):
        try:
            temporary.replace(path)
            return
        except PermissionError as error:
            if getattr(error, "winerror", None) not in {5, 32, 33} or attempt == 5:
                raise
            time.sleep(.01 * (attempt + 1))


class LibraryTask:
    def __init__(self, request: dict, *, timeout: float = 600):
        self.base = Path(tempfile.gettempdir()) / "easy_uiauto" / "library-jobs"
        self.folder = self.base / uuid4().hex
        self.folder.mkdir(parents=True)
        _write(self.folder / "request.json", request)
        self.started = time.monotonic()
        self.timeout = timeout
        self.cancelled_at = None
        self.closed = False
        self.forced_error = None
        self.operation = request.get("operation")
        environment = os.environ.copy()
        environment["PYTHONIOENCODING"] = "utf-8"
        environment["PYTHONUNBUFFERED"] = "1"
        source_root = str(Path(__file__).resolve().parent.parent)
        environment["PYTHONPATH"] = os.pathsep.join(
            filter(None, [source_root, environment.get("PYTHONPATH", "")]))
        self.log = (self.folder / "worker.log").open("wb")
        try:
            command = ([sys.executable, "--worker", str(self.folder)] if getattr(sys,"frozen",False)
                       else [sys.executable, "-m", "easy_uiauto.library_tasks", "--worker", str(self.folder)])
            self.process = subprocess.Popen(
                command,
                env=environment, stdin=subprocess.DEVNULL, stdout=self.log, stderr=self.log,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except BaseException:
            self.log.close()
            self._remove_folder()
            raise

    def cancel(self):
        if self.process.poll() is None and self.cancelled_at is None:
            self.cancelled_at = time.monotonic()
            (self.folder / "cancel").touch()

    def poll(self) -> dict:
        elapsed = time.monotonic() - self.started
        live = self.process.poll() is None
        if live and self.cancelled_at is not None and time.monotonic()-self.cancelled_at >= 2:
            self.process.terminate()
        elif live and elapsed > self.timeout:
            self.forced_error = "任务超过等待预算，已停止本次工作进程；已保存的知识记录仍保留。"
            self.process.terminate()
        try:
            state = json.loads((self.folder / "state.json").read_text(encoding="utf-8"))
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            state = {"state": "running", "message": "正在启动学习工作进程"}
        live = self.process.poll() is None
        if not live and state.get("state") == "running":
            state = {"state": "cancelled" if self.cancelled_at is not None else "failed",
                     "error": self.forced_error or "工作进程提前退出，请查看本次任务日志"}
        if not live and self.forced_error:
            state = {"state": "failed", "error": self.forced_error}
            try:
                saved = json.loads((self.folder / "state.json").read_text(encoding="utf-8"))
                if self.operation == "partition" and self.cancelled_at is None and saved.get("partial_result"):
                    state = {"state": "completed", "result": {**saved["partial_result"],
                             "warning": "UIA 分区读取超时，已停止读取进程；当前为本地像素分区草案，请人工检查。"}}
            except (OSError, ValueError):
                pass
        if self.operation == "pick-region" and self.cancelled_at is not None:
            state = {"state": "cancelled", "error": "已取消区域选择，原区域记录保持不变。"}
        if not live and state.get("state") in {"failed", "cancelled"}:
            try:
                execution = json.loads((self.folder / "execution.json").read_text(encoding="utf-8"))
                state["execution"] = execution
                if execution.get("input_attempted"):
                    completed = sum(s.get("completed", False) for s in execution.get("steps", []))
                    state["error"] = state.get("error", "任务已停止") + f"；已有 {completed} 步完成，已尝试的输入不会自动重试，请检查软件当前状态。"
            except (FileNotFoundError, OSError, json.JSONDecodeError):
                pass
        return {**state, "elapsed_seconds": round(elapsed, 1), "worker_running": live}

    def _remove_folder(self):
        target, intended = self.folder.resolve(), self.base.resolve()
        if target.parent != intended or len(target.name) != 32:
            raise ValueError("Invalid disposable library job path")
        if self.operation == "download-image-models":
            from .image_models import model_specs
            allowed = {Path(spec["path"]).resolve() for spec in model_specs()}
            try:
                paths = json.loads((self.folder / "download-files.json").read_text(encoding="utf-8"))
            except (OSError, ValueError):
                paths = []
            for name in paths:
                path = Path(name).resolve()
                # Only remove a recorded UUID temporary sibling of a known model.
                parts = path.name.rsplit(".", 2)
                if (len(parts) == 3 and len(parts[1]) == 32 and parts[2] == "partial"
                        and all(c in "0123456789abcdef" for c in parts[1])
                        and path.with_name(parts[0]) in allowed):
                    try: path.unlink(missing_ok=True)
                    except OSError: pass
        shutil.rmtree(target, ignore_errors=True)

    def close(self):
        if self.closed:
            return
        if self.process.poll() is None:
            self.process.terminate()
            try:
                self.process.wait(timeout=2)
            except subprocess.TimeoutExpired:
                self.process.kill()
                self.process.wait(timeout=2)
        self.log.close()
        self._remove_folder()
        self.closed = True


@contextmanager
def _capture_guard(folder, enabled):
    if not enabled:
        yield; return
    from .mcp import scanner
    original = scanner._screenshot_window
    def capture(bounds):
        nonce = uuid4().hex
        request, acknowledgement = folder / "capture-request", folder / "capture-ack"
        request.write_text(nonce, encoding="ascii")
        try:
            deadline = time.monotonic() + 3
            while True:
                if (folder / "cancel").exists(): raise InterruptedError("用户停止了当前任务")
                try:
                    if acknowledgement.read_text(encoding="ascii") == nonce: break
                except OSError: pass
                if time.monotonic() >= deadline: raise TimeoutError("无法隐藏任务面板，已停止截图以免污染学习图片")
                time.sleep(.02)
            return original(bounds)
        finally:
            request.unlink(missing_ok=True)
    scanner._screenshot_window = capture
    try: yield
    finally: scanner._screenshot_window = original


def _worker(folder: Path):
    request = json.loads((folder / "request.json").read_text(encoding="utf-8"))
    if request.get("capture_handshake"):
        deadline=time.monotonic()+5
        while not (folder/"gui-ready").exists():
            if (folder/"cancel").exists():
                _write(folder/"state.json",dict(state="cancelled",error="用户停止了当前任务"));return
            if time.monotonic()>=deadline:
                _write(folder/"state.json",dict(state="failed",error="管理器尚未隐藏，未激活或操作目标窗口"));return
            time.sleep(.02)
    with _capture_guard(folder, request.get("capture_handshake", False)):
        _run_worker(folder)


def _run_worker(folder: Path):
    started = time.monotonic()
    partial_result = None

    def progress(message):
        if (folder / "cancel").exists():
            raise InterruptedError("用户停止了本次学习")
        _write(folder / "state.json", {"state": "running", "message": message,
                                       "elapsed_seconds": round(time.monotonic()-started, 1),
                                       **({"partial_result": partial_result} if partial_result else {})})

    try:
        request = json.loads((folder / "request.json").read_text(encoding="utf-8"))
        operation = request["operation"]
        progress("正在检查标注模型" if operation in {"check-image-models", "download-image-models"}
                 else "正在连接所选软件")
        if operation in {"check-image-models", "download-image-models"}:
            from . import image_models
            downloads = []
            def register_download(path):
                downloads.append(str(path))
                _write(folder / "download-files.json", downloads)
            progress.register_download = register_download
            method = (image_models.download_missing_models if operation == "download-image-models"
                      else image_models.inspect_models)
            result = method(progress=progress)
        elif operation in {'inspect-popups','focus-popup','close-popup'}:
            from . import popup_guard
            if operation=='inspect-popups':result=popup_guard.inspect(request['window'])
            else:
                method=popup_guard.close if operation=='close-popup' else popup_guard.focus
                result=method(request['window'],request['popup'])
        elif operation == "import-reference":
            from .knowledge_library import KnowledgeLibrary
            library=KnowledgeLibrary(request["vault"])
            imported=[]
            for source in request["sources"]:
                imported.append(library.import_reference(request["app_id"],source,
                                force=request.get("force",False),progress=progress))
            result=dict(ok=True,reference_ids=[r["id"] for r in imported],
                        message=f"已整理 {len(imported)} 份参考资料；后续学习将读取标准 Markdown")
        elif operation == "supplement-descriptions":
            from .knowledge_library import KnowledgeLibrary
            from . import reference_library
            library=KnowledgeLibrary(request["vault"])
            result=reference_library.supplement(library.directory(request["app_id"]),
                control_ids=request.get("control_ids"),progress=progress)
        elif operation == "open-region":
            import uiautomation
            from .knowledge_library import KnowledgeLibrary
            from .region_actions import open_region
            library=KnowledgeLibrary(request["vault"],page_id=request.get("page_id","manual-main"))
            region=next(r for r in library.read(request["app_id"])["regions"] if r["id"]==request["region_id"])
            with uiautomation.UIAutomationInitializerInThread():
                result=open_region(library,request["app_id"],region,request["window"],progress=progress,
                                   checkpoint=lambda value:_write(folder/"execution.json",value))
        elif operation == "record-steps":
            from .entry_recording import record_steps
            result=record_steps(request["window"],folder,progress)
        elif operation == "pick-region":
            from .region_picker import pick_region

            result = pick_region(request["window"], folder, progress,
                preview=(lambda value: _write(folder / "preview.json", value)) if request.get("picker_preview") else None,
                match_mode=request.get("match_mode","position"))
        elif operation == "capture":
            from .mcp import scanner
            from .ctrl import Controller
            from .mcp import native_windows

            expected = request["window"]
            candidates = native_windows.matching_windows("", handle=expected["handle"])
            if (len(candidates) != 1 or not candidates[0]["visible"]
                    or candidates[0]["process_id"] != expected["process_id"]
                    or candidates[0]["class_name"] != expected["class_name"]):
                raise ValueError("所选软件窗口已经失效，请重新选择")
            window = candidates[0]
            from .popup_guard import assert_unblocked
            assert_unblocked(expected)
            Controller.activate_window_handle(window["handle"])
            bounds = native_windows.visible_window_bounds(
                window["handle"], window["bounds"], class_name=window["class_name"])
            if bounds["width"] < 2 or bounds["height"] < 2:
                raise ValueError("软件没有可见截图区域")
            progress("正在获取软件截图")
            image = scanner._screenshot_window(bounds)
            path = folder / "capture.png"
            image.save(path)
            result = {"image": str(path), "width": image.width, "height": image.height,
                      "window_name": window["title"], "window_class": window["class_name"]}
        elif operation == "partition":
            from PIL import Image
            import uiautomation
            from .library_actions import bind_window
            from .mcp import automatic_layout, native_windows
            from .mcp.control_tree import TreeCollector

            with Image.open(request["image"]) as source:
                image = source.convert("RGB")
            partial_result = automatic_layout.propose(image)
            progress("正在读取窗口容器，生成可调整的区域草案")
            collector = TreeCollector(max_depth=7, max_nodes=600, deadline=time.monotonic()+8,
                                      checkpoint=lambda _state: progress("正在读取窗口容器"))
            warning = ""
            try:
                with uiautomation.UIAutomationInitializerInThread():
                    window = bind_window(request["window"])
                    bounds = native_windows.visible_window_bounds(request["window"]["handle"],
                        native_windows.matching_windows("", handle=request["window"]["handle"])[0]["bounds"],
                        class_name=request["window"]["class_name"])
                    if (bounds["width"], bounds["height"]) != image.size:
                        raise ValueError("窗口尺寸已改变，请刷新截图")
                    collector.visit(window)
                    result = automatic_layout.propose(image, collector.tree,
                        origin=(bounds["left"], bounds["top"]), inspection=collector.state()["inspection"])
            except InterruptedError:
                raise
            except Exception as error:
                warning = f"UIA 分区信息不完整（{type(error).__name__}），当前使用本地像素分区草案，请人工检查。"
                result = partial_result
            if warning: result["warning"] = warning
        elif operation == "workflow":
            import uiautomation
            from .knowledge_library import KnowledgeLibrary
            from . import workflows
            from .mcp import control_observation

            library = KnowledgeLibrary(request["vault"], page_id=request.get("page_id","manual-main"))
            app_id, flow_id = request["app_id"], request["flow_id"]
            index = library.read(app_id)
            flow = next((r for r in index["workflows"] if r["id"] == flow_id), None)
            if flow is None: raise ValueError("流程已不存在")
            with uiautomation.UIAutomationInitializerInThread():
                try:
                    result = workflows.execute(flow, index["controls"], request["window"], parameters=request.get("parameters", {}),
                        directory=library.directory(app_id),
                        progress=progress, checkpoint=lambda r: _write(folder / "execution.json", r),
                        before_action=lambda r, c, main: library.capture_before_click(app_id, r["id"], c, main),
                        after_action=lambda _r, main: control_observation.capture_after(library.directory(app_id), main),
                        after_check=lambda r,c,main:{**(library.capture_before_click(app_id,r["id"],c,main) if c is not None
                            else control_observation.capture_after(library.directory(app_id),main)),"phase":"after-check"})
                except Exception as error:
                    execution = getattr(error, "execution_result", {})
                    if execution:
                        try: library.save_workflow_execution(app_id, flow_id, execution)
                        except Exception as storage_error: error.add_note(str(storage_error))
                    raise
                library.save_workflow_execution(app_id, flow_id, result)
        elif operation in {"learn", "expand", "click"}:
            import uiautomation
            from . import __version__
            from .mcp import configuration, scanner
            from .knowledge_library import KnowledgeLibrary
            from .library_actions import bind_window, execute_plan, owned_windows
            from .mcp import control_observation

            library = KnowledgeLibrary(request["vault"],page_id=request.get("layout",{}).get("page_id",request.get("page_id","manual-main")))
            context = None

            if operation == "click" and request.get("include_region_entry"):
                from . import workflows
                record=next(r for r in library.read(request["app_id"])["controls"] if r["id"]==request["control_id"])
                with uiautomation.UIAutomationInitializerInThread():
                    route=library.entry_steps(request["app_id"],record)
                    if "entry_steps" not in record:
                        from .region_actions import condition_matches
                        state=library.record_state(request["app_id"],record)
                        if state.get("condition") and condition_matches(state["condition"],request["window"],bind_window(request["window"])):
                            route=route[len(state.get("steps",[])):]
                    steps=[*route,workflows.click_step(record,action=request.get("action","click"))]
                    if request.get("text") is not None:steps[-1]["text"]=request["text"]
                    result=workflows.execute(dict(keyword="执行控件",parameters={},steps=steps),library.read(request["app_id"])["controls"],
                        request["window"],directory=library.directory(request["app_id"]),progress=progress,checkpoint=lambda value:_write(folder/"execution.json",value),
                        before_action=lambda r,c,main:library.capture_before_click(request["app_id"],r["id"],c,main),
                        after_action=lambda _r,main:control_observation.capture_after(library.directory(request["app_id"]),main),skip_open_entries=True)
                library.save_execution(request["app_id"],record["id"],result)
                _write(folder/"state.json",{"state":"completed","result":result});return

            if operation == "learn" and request.get("region_id"):
                from .region_actions import open_region
                region=next(r for r in library.read(request["layout"]["app_id"])["regions"] if r["id"]==request["region_id"])
                with uiautomation.UIAutomationInitializerInThread():
                    opened=open_region(library,request["layout"]["app_id"],region,request["window"],progress=progress,
                                       checkpoint=lambda value:_write(folder/"execution.json",value))
                target=opened["target_window"]
                request["layout"].update(window_handle=target["handle"],window_process_id=target["process_id"],window_class=target["class_name"])
                request["window_name"]=target["title"]

            def checkpoint(value):
                _write(folder / "execution.json", value)

            if operation in {"expand", "click"}:
                plan = library.click_plan(request["app_id"], request["control_id"])
                with uiautomation.UIAutomationInitializerInThread():
                    before_ids = {r["id"] for r in library.read(request["app_id"])["controls"]} if operation == "expand" else set()
                    before_popups = {w["handle"] for w in owned_windows(request["window"])} if operation == "expand" else set()
                    try:
                        execution = execute_plan(plan, request["window"], progress=progress, checkpoint=checkpoint,
                            directory=library.directory(request["app_id"]),
                            before_click=lambda record, control, main: library.capture_before_click(
                                request["app_id"], record["id"], control, main,
                                update_reference=operation == "expand"),
                            after_click=lambda _record, main: control_observation.capture_after(
                                library.directory(request["app_id"]), main))
                    except Exception as error:
                        execution = getattr(error, "execution_result", {})
                        try:
                            library.save_execution(request["app_id"], request["control_id"], execution)
                        except Exception as storage_error:
                            error.add_note(f"Saving execution evidence also failed: {storage_error}")
                        raise
                    library.save_execution(request["app_id"], request["control_id"], execution)
                    if operation == "expand":
                        progress("等待展开的控件出现")
                        time.sleep(.2)
                        popups = [w for w in owned_windows(request["window"]) if w["handle"] not in before_popups]
                        context = {"path": [r["id"] for r in plan], "before_ids": before_ids,
                                   "popup_handles": [w["handle"] for w in popups]}
                if operation == "click":
                    progress("点击任务完成")
                    _write(folder / "state.json", {"state": "completed", "result": execution})
                    return

            mode = request.get("semantic_mode", "native")
            settings = configuration.load_gui_configuration() if mode == "api" else {}
            url = ""
            if mode == "api":
                if not all(settings.get(k) for k in ("api_url", "api_key", "api_model")):
                    raise ValueError("请先配置含义补充服务，或关闭“补充控件含义”")
                url, _models = configuration.api_endpoints(settings["api_url"])
            with uiautomation.UIAutomationInitializerInThread():
                result = scanner.scan_window(
                    request["window_name"], url, settings.get("api_key", ""),
                    settings.get("api_model", ""), __version__,
                    root=Path(request["vault"]), max_controls=int(request.get("max_controls", 3000)),
                    strategy="local-visual", manual_layout=request["layout"], semantic_mode=mode,
                    progress=progress, show_overlay=False, navigation_context=context)
            if operation == "expand":
                learned = [r for r in library.read(request["app_id"])["controls"]
                           if r.get("navigation_path") == context["path"]
                           and r["id"] in result["observed_control_ids"]]
                result["navigation_controls"] = len(learned)
                result["execution"] = execution
            if result.get("reference_description_error"):
                result["message"]="基础扫描已完成；参考说明未生成："+result["reference_description_error"]
            elif result.get("reference_descriptions",{}).get("state")=="completed":
                report=result["reference_descriptions"]
                result["message"]=f"学习完成；已生成 {report['matched']} 个控件说明，{len(report['unmatched'])} 个需进一步确认"
        elif operation == "verify":
            import uiautomation
            from .utils import observe_control
            from .knowledge_library import KnowledgeLibrary

            with uiautomation.UIAutomationInitializerInThread():
                if request.get("window"):
                    from .library_actions import bind_window, resolve_record

                    main_window = bind_window(request["window"])
                    control = resolve_record(request, request["window"], main_window,
                                             directory=KnowledgeLibrary(request["vault"]).directory(request["app_id"]))
                else:
                    control = observe_control(request["location"])
                from .locator_images import ImageTarget,target_evidence
                if isinstance(control,ImageTarget):
                    result=dict(ok=True,name=control.name,control_type="",**target_evidence(control),
                                message=f"图片匹配定位成功（分数 {control.score:.3f}，唯一性 {control.uniqueness:.3f}）；未点击控件")
                else:
                    rectangle = control.BoundingRectangle
                    result = {"ok": True, "name": control.Name, "control_type": control.ControlTypeName,
                              "resolved_by":"location", "rect": dict(left=rectangle.left, top=rectangle.top,
                                           right=rectangle.right, bottom=rectangle.bottom),
                              "message": "UIA 定位成功；未点击控件，未改变验证或隔离状态"}
        else:
            raise ValueError("未知知识库任务")
        progress("任务完成")
        _write(folder / "state.json", {"state": "completed", "result": result})
    except InterruptedError as error:
        _write(folder / "state.json", {"state": "cancelled", "error": str(error),
                                       "execution": getattr(error, "execution_result", {})})
    except Exception as error:
        diagnostic=error
        try:
            from .library_actions import diagnose_uia_error
            diagnostic=diagnose_uia_error(error,expected=request.get("window"))
        except Exception:
            pass  # Keep publishing the original failure if UIA/diagnostics cannot load.
        _write(folder / "state.json", {"state": "failed", "error": f"{type(diagnostic).__name__}: {diagnostic}",
                                       "error_code":getattr(diagnostic,"code",""),
                                       "error_reason":getattr(diagnostic,"reason",""),
                                       "popup_state":getattr(diagnostic,"popup_state",{}),
                                       "execution": getattr(error, "execution_result", {})})


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--worker", type=Path, required=True)
    arguments = parser.parse_args()
    _worker(arguments.worker)


if __name__ == "__main__":
    main()
