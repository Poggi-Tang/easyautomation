"""Scoped entrance recording through the existing RecordThread implementation."""
from __future__ import annotations

import ctypes
from ctypes import wintypes
import time


def record_steps(expected, folder, progress):
    import tkinter as tk
    from tkinter import ttk
    import uiautomation as auto
    from .record import RecordThread
    from .ctrl import Controller
    from .library_actions import activate_application, bind_window, owned_windows
    from .region_picker import _native, selection_metadata
    from .mcp import control_observation

    user=_native();allowed={expected["handle"]};cancelled=False
    with auto.UIAutomationInitializerInThread():
        main=bind_window(expected);activate_application(expected)
    class Recorder(RecordThread):
        latest=None
        def track_control(self,control):self.latest=control;super().track_control(control)
        def on_click(self,x,y,button,pressed):
            handle=user.WindowFromPoint(wintypes.POINT(x,y));root=user.GetAncestor(handle,2)
            if int(root or 0) not in allowed:return
            with auto.UIAutomationInitializerInThread():
                super().on_click(x,y,button,pressed)
                if pressed and self.latest is not None and self.actions_data:
                    action=self.actions_data[-1]
                    try:
                        host=bind_window(expected)
                        observation,crop=control_observation.capture(folder,self.latest,host,phase="before-action")
                        scope=({"kind":"main"} if int(root)==expected["handle"] else next(
                            {"kind":"owned","class_name":w["class_name"],"title":w["title"]}
                            for w in owned_windows(expected) if w["handle"]==int(root)))
                        action["capture"]={"selection":{**selection_metadata(self.latest),"location":action["LOCATION"],
                            "normalized_rect":observation["normalized_rect"],"window_scope":scope},
                            "capture":{"image":str(folder/observation["context_image"]),**observation["context_size"],
                                "window_name":expected["title"],"window_class":expected["class_name"],"window_scope":scope}}
                    except Exception as error:action["capture_error"]=str(error)
        def on_press(self,key):
            if key==auto_key_escape or int(auto.GetForegroundWindow() or 0) in allowed:super().on_press(key)
        def on_release(self,key):
            if int(auto.GetForegroundWindow() or 0) in allowed:super().on_release(key)
    from pynput.keyboard import Key
    auto_key_escape=Key.esc
    recorder=Recorder();root=tk.Tk();root.title("easy-uiauto · 录制前置步骤")
    root.attributes("-topmost",True);root.resizable(False,False);root.geometry("420x150+1000+40")
    frame=ttk.Frame(root,padding=12);frame.pack(fill="both",expand=True)
    count=tk.StringVar(value="操作目标软件，按顺序记录前置步骤。")
    ttk.Label(frame,textvariable=count,wraplength=390).pack(anchor="w",pady=6)
    ttk.Button(frame,text="完成录制",command=recorder.stop).pack(side="left",padx=5,pady=10)
    def cancel():
        nonlocal cancelled
        cancelled=True;recorder.stop()
    ttk.Button(frame,text="取消",command=cancel).pack(side="left",padx=5,pady=10)
    root.protocol("WM_DELETE_WINDOW",cancel)
    root.update_idletasks()
    hwnd=user.GetAncestor(root.winfo_id(),2)
    user.SetWindowLongW(hwnd,-20,user.GetWindowLongW(hwnd,-20)|0x08000000|0x80)
    recorder.start();last_inventory=0;ready=False
    def pump():
        nonlocal last_inventory,ready
        if (folder/"cancel").exists():cancel()
        if time.monotonic()-last_inventory>.5:
            allowed.clear();allowed.update({expected["handle"],*(w["handle"] for w in owned_windows(expected))});last_inventory=time.monotonic()
        count.set(f"已记录 {len(recorder.actions_data)} 步；完成后可调整顺序。")
        if recorder.started.is_set() and not ready and recorder.run_error is None:
            with auto.UIAutomationInitializerInThread():activate_application(expected)
            ready=True;progress("正在录制前置步骤；点击浮窗“完成录制”结束，Esc 也可结束录制")
        if not recorder.is_alive():root.quit();return
        root.after(100,pump)
    try:
        progress("正在启动录制监听器…")
        root.after(100,pump);root.mainloop()
    finally:
        recorder.stop();recorder.join(timeout=3);root.destroy()
    if cancelled:raise InterruptedError("已取消录制")
    if recorder.run_error:raise recorder.run_error
    return {"actions":recorder.actions_data}


def archive_actions(library,app_id,actions):
    """Canonicalize recorded actions, retaining repeated clicks and exact order."""
    from . import workflows
    from .mcp.protocol import normalize_location
    from .mcp import knowledge
    names={"点击":"click","双击":"double-click","右击":"right-click","输入文本":"set-text",
           "键盘点击":"press-key","组合键":"hotkey"}
    result=[];last=None
    # Reject unsupported actions before writing any knowledge record.
    for action in actions:
        if action["ACTION"] not in names:raise ValueError("暂不支持此录制动作："+action["ACTION"])
        location=normalize_location(action["LOCATION"])
        if not location["Xpath"] and last is None:raise ValueError("请先点击目标控件，再录制输入或按键")
        if location["Xpath"]:last=True
    last=None
    for action in actions:
        location=normalize_location(action["LOCATION"]);params=location["PARAMETERS"]
        if location["Xpath"]:
            if action.get("capture"):last=library.save_captured_control(app_id,action["capture"])
            else:
                identifier=knowledge.stable_id("control",app_id,library.page_id,location)
                last=dict(id=identifier,app_id=app_id,page_id=library.page_id,location=location,
                          name=location["Name"],semantic_name=location["Name"],control_type=location["ControlType"],
                          class_name=location["ClassName"],automation_id=location["AutomationId"],status="observed",region_id="unassigned")
                knowledge.save_control(library.directory(app_id),last)
        step=workflows.click_step(last,action=names[action["ACTION"]])
        step["text"]=params.get("输入文本",params.get("键盘按键",params.get("组合键","")))
        result.append(step)
    library._refresh(library.directory(app_id));return result
