"""Exercise Tab states and cross-region project creation on an isolated Qt fixture."""
from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import json
from pathlib import Path
from queue import Queue
import subprocess
import sys
import threading
import time
from types import SimpleNamespace

import uiautomation as auto
from PySide6.QtWidgets import QApplication

from easy_uiauto import workflows
from easy_uiauto.ctrl import Controller
from easy_uiauto.knowledge_library import KnowledgeLibrary
from easy_uiauto.library_actions import owned_windows
from easy_uiauto.mcp import native_windows,scanner
from easy_uiauto.mcp.protocol import location_from_xpath
from easy_uiauto.region_picker import normalized_selection,selection_metadata
from easy_uiauto.utils import get_control_xpath
from easy_uiauto.qt_ui.window import KnowledgeWindow


def main():
    parser=argparse.ArgumentParser();parser.add_argument("--output",type=Path,required=True)
    parser.add_argument("--worker-exe",type=Path);args=parser.parse_args()
    if args.worker_exe:
        import easy_uiauto.library_tasks as tasks
        tasks.sys=SimpleNamespace(executable=str(args.worker_exe.resolve()),frozen=True)
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=True)
    if (output/"vault/applications").exists():raise ValueError("Use a fresh verification folder")
    fixture=subprocess.Popen([sys.executable,str(Path(__file__).resolve().parents[1]/"tests/fixtures/qt_knowledge_target.py"),
                              "--events",str(output/"events.json")],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding="utf-8")
    gui=None;target=None;app=QApplication.instance() or QApplication([])
    from easy_uiauto.qt_gui import configure_application
    configure_application(app)
    try:
        queue=Queue();threading.Thread(target=lambda:queue.put(fixture.stdout.readline()),daemon=True).start()
        line=queue.get(timeout=20);assert line.startswith("fixture_ready hwnd="),line
        target=native_windows.matching_windows("",handle=int(line.split("=",1)[1]))[0]
        root=auto.ControlFromHandle(target["handle"]);Controller.activate_window_handle(target["handle"],root)
        library=KnowledgeLibrary(output/"vault");library.bind("qt-demo","Qt 验收软件",target)

        def wait(predicate,seconds=20):
            deadline=time.monotonic()+seconds
            while time.monotonic()<deadline:
                app.processEvents()
                value=predicate()
                if value:return value
                time.sleep(.05)
            raise AssertionError("Acceptance wait timed out")

        def find(name,scope=None,kind=None):
            roots=[scope] if scope is not None else [root,*[auto.ControlFromHandle(w["handle"]) for w in owned_windows(target)]]
            matches=[]
            for parent in roots:
                for control,_depth in scanner._walk_controls(parent,16,400):
                    if control.Name==name and not control.IsOffscreen and (not kind or control.ControlTypeName==kind):
                        if not any(auto.ControlsAreSame(control,c) for c in matches):matches.append(control)
            if len(matches)!=1:
                tree=[(c.ControlTypeName,c.Name) for parent in roots for c,_ in scanner._walk_controls(parent,12,150)]
                raise AssertionError(dict(name=name,matches=len(matches),tree=tree))
            return matches[0]

        def picked(control,scope=None):
            bounds=scanner._window_rect(root)
            capture=output/("capture-"+str(time.time_ns())+".png");scanner._screenshot_window(bounds).save(capture)
            metadata=selection_metadata(control)
            return dict(selection={**metadata,"location":location_from_xpath(get_control_xpath(control,match_mode="identity")),
                         "normalized_rect":normalized_selection(metadata["rect"],bounds),"window_scope":scope or {"kind":"main"}},
                        capture=dict(image=str(capture),width=bounds["width"],height=bounds["height"],
                                     window_name=target["title"],window_class=target["class_name"]))

        def capture(control,scope=None):return library.save_captured_control("qt-demo",picked(control,scope))

        home_tab=capture(find("主页",kind="TabItemControl"));plot_tab=capture(find("绘图",kind="TabItemControl"))
        base_id=library.save_picked_region("qt-demo",picked(find("主页区域")),name="工作区")
        base=next(r for r in library.regions("qt-demo") if r["id"]==base_id)
        library.configure_region("qt-demo",base_id,steps=[workflows.click_step(home_tab)],state_name="主页",
            condition=dict(location=home_tab["location"],kind="selected",label="主页"))
        old_image=library.asset("qt-demo",base["image"]).read_bytes()
        Controller.execute_resolved("click",control=find("绘图",kind="TabItemControl"));wait(lambda:find("绘图区域"))
        variant_id=library.save_variant("qt-demo",base,picked(find("绘图区域")),name="绘图",steps=[workflows.click_step(plot_tab)],
            condition=dict(location=plot_tab["location"],kind="selected",label="绘图"))
        assert library.asset("qt-demo",base["image"]).read_bytes()==old_image

        tree_record=capture(find("工程树"));file_record=capture(find("文件"))
        Controller.execute_resolved("click",control=find("文件"));new_record=capture(find("新建工程"))
        Controller.execute_resolved("click",control=find("新建工程"))
        dialog_window=wait(lambda:next((w for w in owned_windows(target) if w["title"]=="新建工程"),None))
        dialog=auto.ControlFromHandle(dialog_window["handle"])
        dialog_scope=dict(kind="owned",class_name=dialog_window["class_name"],title=dialog_window["title"])
        dialog_record=capture(dialog,dialog_scope);name_record=capture(find("工程名称",dialog,kind="EditControl"),dialog_scope)
        confirm_record=capture(find("确认",dialog),dialog_scope)
        Controller.execute_resolved("click",control=find("取消",dialog))
        wait(lambda:not any(w["title"]=="新建工程" for w in owned_windows(target)))

        pre=workflows.click_step(tree_record,action="inspect")
        pre["checks"]=[dict(id="before-node",kind="node-absent",target=dict(control_id=tree_record["id"],label="工程树"),expected="${工程名称}")]
        file_step=workflows.click_step(file_record);new_step=workflows.click_step(new_record)
        new_step["checks"]=[dict(id="dialog-visible",kind="exists",target=dict(control_id=dialog_record["id"],label="新建工程"))]
        text_step=workflows.click_step(name_record,action="set-text");text_step["text"]="${工程名称}"
        text_step["checks"]=[dict(id="name-equals",kind="text",target=dict(control_id=name_record["id"],label="工程名称"),expected="${工程名称}")]
        confirm=workflows.click_step(confirm_record)
        confirm["checks"]=[dict(id="dialog-closed",kind="not-visible",target=dict(control_id=dialog_record["id"],label="新建工程")),
                           dict(id="after-node",kind="node-exists",target=dict(control_id=tree_record["id"],label="工程树"),expected="${工程名称}")]
        flow=library.save_workflow("qt-demo",dict(id="new-project",keyword="新建工程",parameters={"工程名称":"Qt_Demo_001"},
            precondition="软件已启动",steps=[pre,file_step,new_step,text_step,confirm]))
        assert flow["status"]=="ready",flow["readiness"]

        gui=KnowledgeWindow(library.root,remember=False);gui.windows["qt-demo"]=target
        errors=[];gui.error=lambda title,error:errors.append(f"{title}: {error}")
        gui.show();wait(lambda:"qt-demo" in gui.snapshots)
        gui.goto("region:qt-demo:"+base_id);gui.learn();wait(lambda:gui.task is None,90);assert not errors,errors
        wait(lambda:any(c.get("page_id")=="manual-main" and c.get("region_id")==base_id.split(".",1)[1] for c in gui.snapshots["qt-demo"]["controls"]))
        home_ids={c["id"] for c in library.read("qt-demo")["controls"] if c.get("region_id")==base_id.split(".",1)[1]}
        gui.goto("region:qt-demo:"+variant_id);app.processEvents();gui.grab().save(str(output/"regions.png"))
        gui.learn();wait(lambda:gui.task is None,90);assert not errors,errors
        wait(lambda:any(c.get("page_id")==variant_id.split(".",1)[0] for c in gui.snapshots["qt-demo"]["controls"]))
        assert library.asset("qt-demo",base["image"]).read_bytes()==old_image
        assert all(c.get("status")!="quarantined" for c in library.read("qt-demo")["controls"] if c["id"] in home_ids)
        plot_ids={c["id"] for c in library.read("qt-demo")["controls"] if c.get("page_id")==variant_id.split(".",1)[0]}
        gui.goto("region:qt-demo:"+base_id);gui.learn();wait(lambda:gui.task is None,90);assert not errors,errors
        assert plot_ids and not (home_ids & plot_ids)
        assert all(c.get("status")!="quarantined" for c in library.read("qt-demo")["controls"] if c["id"] in plot_ids)
        gui.goto("workflow:qt-demo:new-project");app.processEvents();gui.grab().save(str(output/"steps.png"))
        gui.start_task("workflow",dict(vault=str(library.root),app_id="qt-demo",flow_id=flow["id"],window=target,parameters={"工程名称":"Qt_Demo_001"}),gui.display_result,hide=True)
        wait(lambda:gui.task is None,90);assert not errors,errors
        assert gui.result.get("ok"),gui.result
        checks=[c for s in gui.result["steps"] for c in s["checks"]]
        assert checks and all(c.get("observation") for c in checks),checks
        events=json.loads((output/"events.json").read_text(encoding="utf-8"))
        assert sum(e["event"]=="confirm" for e in events)==1,events
        app.processEvents();gui.grab().save(str(output/"results.png"))
        report=dict(ok=True,states_separate=True,old_image_preserved=True,region_learning=True,
                    cross_region_checks=True,confirm_count=1,run=gui.result,errors=errors)
        (output/"acceptance.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
        print(json.dumps({k:v for k,v in report.items() if k!="run"},ensure_ascii=True),flush=True)
    finally:
        if gui:
            if gui.task:gui.task.close();gui.task=None
            gui.dirty=False;gui.close();app.processEvents()
        if target and fixture.poll() is None:
            user=ctypes.WinDLL("user32");user.PostMessageW.argtypes=[wintypes.HWND,wintypes.UINT,wintypes.WPARAM,wintypes.LPARAM]
            user.PostMessageW(target["handle"],0x10,0,0)
            try:fixture.wait(timeout=3)
            except subprocess.TimeoutExpired:fixture.terminate();fixture.wait(timeout=3)
        elif fixture.poll() is None:fixture.terminate();fixture.wait(timeout=3)


if __name__=="__main__":main()
