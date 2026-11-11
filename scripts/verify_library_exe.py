"""Cold-start the compiled manager against a fixture vault and inspect its UIA."""
from pathlib import Path
import argparse
import ctypes
from ctypes import wintypes
import json
import subprocess
import time

import uiautomation as auto
from easy_uiauto.ctrl import Controller
from easy_uiauto.mcp import native_windows,scanner


def main():
    parser=argparse.ArgumentParser();parser.add_argument("--exe",type=Path,required=True)
    parser.add_argument("--vault",type=Path,required=True);parser.add_argument("--output",type=Path,required=True)
    args=parser.parse_args();args.output.mkdir(parents=True,exist_ok=True)
    start=time.monotonic()
    try:process=subprocess.Popen([str(args.exe.resolve()),"--vault",str(args.vault.resolve())])
    except OSError as error:
        if getattr(error,"winerror",None)==740:
            raise SystemExit("EXE 要求管理员权限，请在管理员终端运行此验证脚本。") from error
        raise
    window=None
    try:
        deadline=time.monotonic()+30
        while time.monotonic()<deadline:
            found=[w for w in native_windows.enumerate_windows(visible_only=True)
                   if w["process_id"]==process.pid and w["title"]=="easy-uiauto 知识库管理" and w["class_name"].startswith("Qt")]
            if found:window=found[0];break
            if process.poll() is not None:raise AssertionError(("GUI exited",process.returncode))
            time.sleep(.1)
        assert window,"Frozen Qt main window did not appear"
        startup=time.monotonic()-start;root=auto.ControlFromHandle(window["handle"])
        Controller.activate_window_handle(window["handle"],root)
        names=set();deadline=time.monotonic()+5
        while time.monotonic()<deadline:
            names={str(c.Name) for c,_ in scanner._walk_controls(root,18,500)}
            if {"设置","选择窗口","添加软件","知识库导航"}.issubset(names):break
            time.sleep(.1)
        scanner._screenshot_window(scanner._window_rect(root)).save(args.output/"frozen-main.png")
        tree=[dict(depth=depth,name=str(c.Name),kind=c.ControlTypeName,rect=scanner._rect(c)) for c,depth in scanner._walk_controls(root,18,500)]
        (args.output/"tree.json").write_text(json.dumps(tree,ensure_ascii=False,indent=2),encoding="utf-8")
        assert {"设置","选择窗口","添加软件","知识库导航"}.issubset(names),names
        result=dict(ok=True,main_window=True,accessible_controls=True,startup_seconds=round(startup,2))
        (args.output/"frozen-main.json").write_text(json.dumps(result),encoding="utf-8");print(json.dumps(result),flush=True)
    finally:
        if window:
            user=ctypes.WinDLL("user32");user.PostMessageW.argtypes=[wintypes.HWND,wintypes.UINT,wintypes.WPARAM,wintypes.LPARAM]
            user.PostMessageW(window["handle"],0x10,0,0)
        try:process.wait(timeout=5)
        except subprocess.TimeoutExpired:process.terminate();process.wait(timeout=3)


if __name__=="__main__":main()
