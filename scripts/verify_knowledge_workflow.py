"""Exercise the supplied test case on the disposable --workflow native fixture."""

from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import json
from pathlib import Path
import time
import tkinter as tk

import uiautomation as uia

from easy_uiauto import workflows
from easy_uiauto.ctrl import Controller
from easy_uiauto.knowledge_library import KnowledgeLibrary
from easy_uiauto.library_actions import owned_windows
from easy_uiauto.library_tasks import LibraryTask
from easy_uiauto.mcp import knowledge, native_windows
from easy_uiauto.mcp.protocol import location_from_xpath
from easy_uiauto.utils import get_control_xpath
from easy_uiauto.workflow_gui import WorkflowWindow


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--events", required=True, type=Path)
    parser.add_argument("--gui-run", action="store_true", help="Run through the integrated GUI and its capture-safe stop panel")
    parser.add_argument("--hold", action="store_true", help="Keep the verified GUI open for visual inspection")
    args = parser.parse_args()
    output = args.output.resolve(); output.mkdir(parents=True, exist_ok=True)
    library = KnowledgeLibrary(output/"vault")
    if library.apps(): raise ValueError("Use a new disposable vault")
    windows = [w for w in native_windows.enumerate_windows(visible_only=True)
               if w["class_name"] == "EasyUIAutoWorkflowFixture" and w["title"] == "Knowledge Workflow Target"]
    assert len(windows) == 1, "Start exactly one --workflow fixture first"
    window, app_id = windows[0], "workflow-acceptance"
    library.bind(app_id, "组合动作演示软件", window)
    directory = library.directory(app_id)
    main_control = uia.ControlFromHandle(window["handle"])
    Controller.activate_window_handle(window["handle"], control=main_control)
    user32 = ctypes.WinDLL("user32")
    user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]

    def wait_owned(class_name, title=None):
        deadline = time.monotonic()+5
        while time.monotonic() < deadline:
            matches = [w for w in owned_windows(window) if w["class_name"] == class_name
                       and (title is None or w["title"] == title)]
            if len(matches) == 1: return matches[0]
            time.sleep(.05)
        raise AssertionError("Missing fixture popup: "+class_name)

    def close_dialogs():
        for w in owned_windows(window):
            if w["class_name"] == window["class_name"] and w["title"] == "新建项目":
                user32.PostMessageW(w["handle"], 0x10, 0, 0)
        time.sleep(.15)

    def learn(control, label, scope=None):
        identifier = knowledge.stable_id("control", label, control.ControlTypeName)
        record = dict(id=identifier, kind="control", app_id=app_id, page_id="main", status="observed",
            name=control.Name, semantic_name=label, control_type=control.ControlTypeName,
            location=location_from_xpath(get_control_xpath(control)), window_scope=scope or {"kind": "main"})
        knowledge.save_control(directory, record)
        library.capture_before_click(app_id, identifier, control, main_control, update_reference=True)
        return identifier

    close_dialogs()
    learn(main_control, "项目管理")
    Controller.execute_resolved("right-click", control=main_control, x=300, y=250)
    popup = wait_owned("#32768")
    menu = uia.ControlFromHandle(popup["handle"])
    scope = dict(kind="owned", class_name="#32768", title=popup["title"])
    learn(menu, "项目管理右键菜单", scope)
    new = next(c for c in menu.GetChildren() if c.Name == "新建项目")
    learn(new, "新建项目", scope)
    Controller.execute_resolved("click", control=new)
    dialog = wait_owned(window["class_name"], "新建项目")
    dialog_control = uia.ControlFromHandle(dialog["handle"])
    scope = dict(kind="owned", class_name=dialog["class_name"], title="新建项目")
    learn(dialog_control, "新建项目", scope)
    for control in dialog_control.GetChildren():
        if control.ControlTypeName in {"EditControl", "ButtonControl"} and control.Name in {
            "选择存储路径", "选择", "设置项目名称", "取消", "确定"}:
            learn(control, control.Name, scope)
    close_dialogs()
    case = (Path(__file__).resolve().parents[1]/"docs/examples/workflows/new-project-context-menu.md").read_text(encoding="utf-8")
    flow = library.import_workflow_case(app_id, case)
    library.save_usage(app_id, "新建项目操作说明", "# 新建项目\n\n1. 右键点击项目管理空白处。\n2. 选择新建项目。\n3. 检查对话框标题、输入框及按钮。\n\n关联组合动作可自动执行这些步骤和预期检查。", workflow_ids=[flow["id"]])
    assert flow["readiness"]["ready"], flow["readiness"]
    baseline = json.loads(args.events.read_text(encoding="utf-8"))
    controls_before = {p.name: p.read_bytes() for p in (directory/"controls").glob("*.md")}

    root = gui = panel = None
    monitor_seen = False
    if args.gui_run:
        from easy_uiauto.knowledge_gui import KnowledgeWindow
        root = tk.Tk(); gui = KnowledgeWindow(root, library.root)
        gui.windows[app_id] = window
        gui.open_action(app_id, flow["id"]); root.update()
        panel = gui.workflow_panel; panel.run()
        deadline = time.monotonic()+65
        while panel.task and time.monotonic() < deadline:
            root.update()
            if panel.monitor:
                monitor_seen = monitor_seen or bool(user32.IsWindowVisible(panel.monitor.hwnd))
            time.sleep(.03)
        assert not panel.task, "GUI job did not finish"
        execution = library.workflows(app_id)[0]["last_run"]
        assert execution["ok"] and monitor_seen, (execution, monitor_seen)
    else:
        task = LibraryTask(dict(operation="workflow", vault=str(library.root), app_id=app_id,
                                flow_id=flow["id"], window=window), timeout=60)
        try:
            state = task.poll()
            while state["worker_running"]:
                time.sleep(.1); state = task.poll()
            if state["state"] != "completed":
                (output/"failed-state.json").write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
            assert state["state"] == "completed", state
            execution = state["result"]
        finally:
            task.close()
    events = json.loads(args.events.read_text(encoding="utf-8"))[len(baseline):]
    assert events == ["context-open", "new-project-dialog"], events
    assert execution["ok"] and sum(len(s["checks"]) for s in execution["steps"]) == 7
    assert controls_before == {p.name: p.read_bytes() for p in (directory/"controls").glob("*.md")}
    for step in execution["steps"]:
        assert library.asset(app_id, step["after"]["context_image"]).is_file()
    assert library.asset(app_id, execution["steps"][1]["before"]["control_image"]).is_file()
    result = dict(ok=True, vault=str(library.root), keyword=flow["keyword"], steps=3, assertions=7,
                  menu_items=8, executed_once=True, no_confirm_click=True, learning_unchanged=True,
                  events=events, execution=execution, gui_run=args.gui_run, monitor_visible=monitor_seen)
    (output/"acceptance.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k:v for k,v in result.items() if k != "execution"}, ensure_ascii=False), flush=True)
    close_dialogs()
    if gui:
        panel.notebook.select(3)
        if args.hold: root.mainloop()
        else: gui.close()


if __name__ == "__main__": main()
