"""Exercise real hover, wheel, Ctrl and Esc only on our disposable native fixture."""
from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import json
from pathlib import Path
import subprocess
import sys
from queue import Queue
import threading
import time
import tkinter as tk

import uiautomation as auto
from PIL import ImageGrab

from easy_uiauto.ctrl import Controller
from easy_uiauto.knowledge_gui import KnowledgeWindow, WindowChooser
from easy_uiauto.library_tasks import LibraryTask
from easy_uiauto.mcp import native_windows
from easy_uiauto.utils import observe_control


def main():
    parser = argparse.ArgumentParser(); parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(); output = args.output.resolve(); output.mkdir(parents=True, exist_ok=True)
    if (output / "vault/applications").exists(): raise ValueError("Use a fresh verification folder")
    root_dir = Path(__file__).resolve().parents[1]
    events_file = output / "events.json"
    fixture = subprocess.Popen([sys.executable, str(root_dir / "tests/fixtures/knowledge_target.py"),
                                "--picker", "--events", str(events_file)], stdout=subprocess.PIPE,
                               stderr=subprocess.STDOUT, text=True, encoding="utf-8")
    root = gui = None
    try:
        startup = Queue()
        threading.Thread(target=lambda: startup.put(fixture.stdout.readline()), daemon=True).start()
        line = startup.get(timeout=15)
        assert line.startswith("fixture_ready hwnd="), line
        handle = int(line.split("=",1)[1])
        deadline = time.monotonic()+15
        target = None
        while time.monotonic() < deadline:
            windows = [w for w in native_windows.enumerate_windows(visible_only=True)
                       if w["handle"] == handle and w["class_name"] == "EasyUIAutoKnowledgeFixture"]
            if len(windows) == 1: target = windows[0]; break
            time.sleep(.05)
        assert target, "Our fixture failed to create its window"
        root = tk.Tk(); gui = KnowledgeWindow(root, output / "vault")
        import easy_uiauto.knowledge_gui as module
        errors = []
        module.messagebox.showerror = lambda title, message, **kw: errors.append(f"{title}: {message}")
        module.LibraryTask = lambda request, **kw: LibraryTask({**request, "picker_preview": True}, **kw)
        root.update()

        def wait_until(predicate, timeout=20):
            limit = time.monotonic()+timeout
            while time.monotonic() < limit:
                root.update()
                if errors: raise AssertionError(errors)
                value = predicate()
                if value: return value
                time.sleep(.04)
            raise AssertionError("Timed out: " + gui.status.get())

        def completed():
            wait_until(lambda: gui.task is None, 25)

        def ready():
            task = gui.task
            def check():
                state = task.poll()
                if state["state"] == "failed": raise AssertionError(state)
                return "Ctrl" in state.get("message", "")
            wait_until(check)
            return task

        def preview(task, **expected):
            def check():
                if gui.task is None: raise AssertionError("Picker finished before confirmation: " + gui.status.get())
                try: value = json.loads((task.folder / "preview.json").read_text(encoding="utf-8"))
                except (OSError, ValueError): return None
                return value if value.get("token") and all(value.get(k) == v for k,v in expected.items()) else None
            return wait_until(check)

        def outline(value, image_name):
            r = value["rect"]; left, top, right, bottom = (r[k] for k in ("left","top","right","bottom"))
            expected = {(left,top,right,top+2), (left,bottom-2,right,bottom),
                        (left,top,left+2,bottom), (right-2,top,right,bottom)}
            def visible():
                bars = [w for w in native_windows.enumerate_windows(visible_only=True)
                        if w["title"] == "easy-uiauto · 添加 / 修改区域"
                        and (w["bounds"]["width"] == 2 or w["bounds"]["height"] == 2)]
                actual = {tuple(w["bounds"][k] for k in ("left","top","right","bottom")) for w in bars}
                return actual == expected
            wait_until(visible, 5)
            image = ImageGrab.grab(bbox=(left-4,top-4,right+4,bottom+4), all_screens=True)
            # Check the screen pixels, not only the existence of four overlay HWNDs.
            for x,y in ((left+(right-left)//2,top), (left+(right-left)//2,bottom-1),
                        (left,top+(bottom-top)//2), (right-1,top+(bottom-top)//2)):
                pixel = image.getpixel((x-left+4,y-top+4))[:3]
                assert pixel == (255,23,68), {"pixel":pixel,"point":(x,y),"rect":r}
            image.save(output / image_name)

        chooser = WindowChooser(root, gui.library, "", gui.bound)
        chooser.app_id.set("picker-acceptance"); chooser.name.set("区域拾取演示软件")
        chooser.tree.selection_set(str(target["handle"])); chooser.accept(); completed()
        assert not gui.library.regions(gui.app_id), "Window selection must not silently create regions"
        main_window = auto.ControlFromHandle(target["handle"])
        button = main_window.ButtonControl(Name="Save", searchDepth=2)

        gui.buttons["添加区域"].invoke(); task = ready()
        Controller.execute_resolved("hover", control=button)
        leaf = preview(task, native_name="Save", control_type="ButtonControl", level=0)
        assert leaf["automation_id"] == "102" and leaf["class_name"] == "Button"
        outline(leaf, "hover-save.png")
        Controller.execute_resolved("hover", control=main_window.ButtonControl(Name="Open", searchDepth=2))
        moved = preview(task, native_name="Open", control_type="ButtonControl", level=0)
        outline(moved, "hover-open.png")
        Controller.execute_resolved("hover", control=button)
        leaf = preview(task, native_name="Save", control_type="ButtonControl", level=0)
        outline(leaf, "hover-save-again.png")
        position = auto.GetCursorPos()
        Controller.execute_resolved("wheel-up", x=position[0], y=position[1])
        parent = preview(task, control_type="WindowControl", level=1)
        outline(parent, "wheel-parent.png")
        Controller.execute_resolved("wheel-down", x=position[0], y=position[1])
        child = preview(task, native_name="Save", level=0)
        outline(child, "wheel-child.png")
        visible_windows = native_windows.enumerate_windows(visible_only=True)
        inspectors = [w for w in visible_windows
                      if w["title"] == "easy-uiauto · 添加 / 修改区域" and w["bounds"]["width"] > 250
                      and w["bounds"]["height"] > 150]
        assert len(inspectors) == 1, {"inspectors": inspectors, "task": task.poll(),
            "windows": [(w["title"],w["class_name"]) for w in visible_windows if w["class_name"] == "TkTopLevel"]}
        assert not inspectors[0]["foreground"], "Inspector stole focus from the target"
        assert next(w for w in visible_windows if w["handle"] == target["handle"])["foreground"]
        b = inspectors[0]["bounds"]
        ImageGrab.grab(bbox=(b["left"],b["top"],b["right"],b["bottom"]), all_screens=True).save(output / "picker.png")
        Controller.hotkey("ctrl"); completed()
        regions = gui.library.regions(gui.app_id); assert len(regions) == 1
        original_id = regions[0]["id"]
        assert auto.ControlsAreSame(observe_control(regions[0]["location"], window=main_window), button)
        assert gui.library.page(gui.app_id)["layout_confirmed"]

        gui.buttons["修改区域"].invoke(); task = ready()
        Controller.execute_resolved("hover", control=button)
        preview(task, native_name="Save", level=0)
        Controller.execute_resolved("wheel-up", x=position[0], y=position[1])
        preview(task, control_type="WindowControl", level=1)
        Controller.hotkey("ctrl"); completed()
        regions = gui.library.regions(gui.app_id)
        assert len(regions) == 1 and regions[0]["id"] == original_id
        assert regions[0]["control_type"] == "WindowControl" and regions[0]["name"] == "Save"
        assert auto.ControlsAreSame(observe_control(regions[0]["location"], window=main_window), main_window)

        before = gui.library.read(gui.app_id)
        gui.windows.clear()  # Simulate restarting the manager with an existing vault.
        gui.buttons["添加区域"].invoke(); ready()
        Controller.hotkey("esc"); completed()
        assert gui.library.read(gui.app_id) == before and "取消" in gui.status.get()
        leaked = json.loads(events_file.read_text(encoding="utf-8"))
        assert leaked == [], f"Picker input leaked to target: {leaked}"
        result = dict(ok=True, hover=leaf, parent=parent, ctrl_saved=True, modification_kept_id=True,
                      visible_outline_after_mouse_move=True, visible_outline_after_parent_child_switch=True,
                      xpath_replayed_to_selected_control=True, cancellation_preserved_records=True,
                      existing_software_rebound=True, target_input_events=leaked, errors=errors)
        (output / "acceptance.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
    finally:
        if gui:
            if gui.task: gui.task.close(); gui.task = None
            gui.dirty = gui.workflow_panel.dirty = gui.usage_panel.dirty = False; gui.close()
        elif root: root.destroy()
        if fixture.poll() is None:
            if target:
                user = ctypes.WinDLL("user32")
                user.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
                user.PostMessageW(target["handle"], 0x10, 0, 0)
                try: fixture.wait(timeout=3)
                except subprocess.TimeoutExpired: pass
        if fixture.poll() is None:
            fixture.terminate(); fixture.wait(timeout=5)


if __name__ == "__main__": main()
