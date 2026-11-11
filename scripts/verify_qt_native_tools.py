"""Verify shared recording and modal-dialog picking on our disposable Qt target."""
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
from easy_uiauto.ctrl import Controller
from easy_uiauto.knowledge_library import KnowledgeLibrary
from easy_uiauto.library_actions import resolve_record,validate_plan,execute_plan
from easy_uiauto.library_tasks import LibraryTask
from easy_uiauto.mcp import native_windows,scanner


def main():
    parser=argparse.ArgumentParser();parser.add_argument("--output",type=Path,required=True)
    parser.add_argument("--worker-exe",type=Path);args=parser.parse_args()
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=True)
    if (output/"vault/applications").exists():raise ValueError("Use a fresh test folder")
    if args.worker_exe:
        import easy_uiauto.library_tasks as tasks
        tasks.sys=SimpleNamespace(executable=str(args.worker_exe.resolve()),frozen=True)
    fixture=subprocess.Popen([sys.executable,str(Path(__file__).resolve().parents[1]/"tests/fixtures/qt_knowledge_target.py"),
                              "--events",str(output/"events.json")],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding="utf-8")
    target=task=None
    def wait(predicate,seconds=25):
        until=time.monotonic()+seconds
        while time.monotonic()<until:
            result=predicate()
            if result:return result
            time.sleep(.05)
        raise AssertionError("Native-tool verification timed out")
    def status():
        state=task.poll()
        if not state["worker_running"] and state["state"] not in {"completed","cancelled"}:raise AssertionError(state)
        return state
    def finish():
        state=wait(lambda:(s if not s["worker_running"] else None) if (s:=status()) else None)
        assert state["state"]=="completed",state
        return state["result"]
    def find(parent,name,kind=None):
        found=[c for c,_ in scanner._walk_controls(parent,16,300) if c.Name==name and not c.IsOffscreen and (not kind or c.ControlTypeName==kind)]
        assert len(found)==1,(name,len(found));return found[0]
    def events():return json.loads((output/"events.json").read_text(encoding="utf-8"))
    try:
        queue=Queue();threading.Thread(target=lambda:queue.put(fixture.stdout.readline()),daemon=True).start()
        line=queue.get(timeout=20);assert line.startswith("fixture_ready hwnd="),line
        target=native_windows.matching_windows("",handle=int(line.split("=",1)[1]))[0]
        root=auto.ControlFromHandle(target["handle"])
        library=KnowledgeLibrary(output/"vault");library.bind("native-tools","Native tools",target)
        task=LibraryTask(dict(operation="record-steps",window=target),timeout=60)
        wait(lambda:"正在录制前置步骤" in status().get("message",""))
        Controller.execute_resolved("click",control=find(root,"主页动作"))
        wait(lambda:any(e["event"]=="action" for e in events()))
        Controller.hotkey("esc");recorded=finish()
        from easy_uiauto.entry_recording import archive_actions
        steps=archive_actions(library,"native-tools",recorded["actions"])
        assert len(steps)==1 and steps[0]["action"]=="click",recorded
        task.close();task=None
        plan=validate_plan(library.read("native-tools")["controls"],steps[0]["target"]["control_id"],path=[])
        execute_plan(plan,target)
        wait(lambda:sum(e["event"]=="action" for e in events())==2)
        Controller.execute_resolved("click",control=find(root,"文件"))
        Controller.execute_resolved("click",control=find(root,"新建工程",kind="MenuItemControl"))
        dialog_meta=wait(lambda:next((w for w in native_windows.enumerate_windows(visible_only=True)
                                     if w["process_id"]==target["process_id"] and w["title"]=="新建工程"),None))
        dialog=auto.ControlFromHandle(dialog_meta["handle"])
        task=LibraryTask(dict(operation="pick-region",window=target,picker_preview=True),timeout=60)
        wait(lambda:"Ctrl" in status().get("message",""))
        Controller.execute_resolved("hover",control=find(dialog,"工程名称",kind="EditControl"))
        def preview():
            status()
            try:r=json.loads((task.folder/"preview.json").read_text(encoding="utf-8"))
            except (OSError,ValueError):return None
            return r if r.get("native_name")=="工程名称" and r.get("control_type")=="EditControl" and r.get("token") else None
        wait(preview);Controller.hotkey("ctrl");picked=finish()
        assert picked["selection"]["window_scope"]["kind"]=="owned",picked
        record=library.save_captured_control("native-tools",picked)
        assert auto.ControlsAreSame(resolve_record(record,target,root),find(dialog,"工程名称",kind="EditControl"))
        task.close();task=None
        Controller.execute_resolved("click",control=find(dialog,"取消"))
        report=dict(ok=True,shared_recording_replayed=True,recorded_actions=len(steps),
                    modal_dialog_pick=True,owned_xpath_replayed=True)
        (output/"acceptance.json").write_text(json.dumps(report,indent=2),encoding="utf-8")
        print(json.dumps(report),flush=True)
    finally:
        if task:task.close()
        if target and fixture.poll() is None:
            user=ctypes.WinDLL("user32");user.PostMessageW.argtypes=[wintypes.HWND,wintypes.UINT,wintypes.WPARAM,wintypes.LPARAM]
            user.PostMessageW(target["handle"],0x10,0,0)
            try:fixture.wait(timeout=3)
            except subprocess.TimeoutExpired:fixture.terminate();fixture.wait(timeout=3)
        elif fixture.poll() is None:fixture.terminate();fixture.wait(timeout=3)


if __name__=="__main__":main()
