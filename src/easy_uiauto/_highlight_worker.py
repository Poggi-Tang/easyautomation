"""Internal JSON-lines worker; the Tk mainloop stays on its process main thread."""

from __future__ import annotations

import json
import queue
import sys
import threading


def _emit(event, **data):
    print(json.dumps({"event": event, **data}, ensure_ascii=False), flush=True)


def main():
    import tkinter as tk

    import uiautomation

    from .draw import ScreenLineBox
    from .utils import format_action_data_by_xpath, observe_control

    root = box = None
    try:
        initial = json.loads(sys.stdin.readline())
        if initial.get("cmd") != "init":
            raise ValueError("first highlight command must be init")
        with uiautomation.UIAutomationInitializerInThread():
            root = tk.Tk()
            root.withdraw()
            box = ScreenLineBox(root=root, interval_ms=50, **(initial.get("style") or {}))
            if initial.get("rect"):
                box.set_rect(initial["rect"])
            commands = queue.Queue()

            def read_commands():
                try:
                    for line in sys.stdin:
                        commands.put(json.loads(line))
                except Exception as error:
                    commands.put({"cmd": "reader_error", "error": str(error)})
                finally:
                    commands.put({"cmd": "stop"})

            threading.Thread(target=read_commands, daemon=True).start()

            def pump():
                for _ in range(100):
                    try:
                        command = commands.get_nowait()
                    except queue.Empty:
                        break
                    try:
                        name = command.get("cmd")
                        if name == "stop":
                            root.quit()
                            return
                        if name == "reader_error":
                            raise RuntimeError(command["error"])
                        if name == "rect":
                            box.set_rect(command.get("rect"))
                        elif name == "set_xpath":
                            action = format_action_data_by_xpath("", "", command["xpath"], {}, "")
                            box.set_control(observe_control(action["LOCATION"]))
                        elif name == "style":
                            box.set_style(**(command.get("style") or {}))
                    except Exception as error:
                        _emit("error", message=f"highlight command failed: {error}")
                root.after(16, pump)

            root.after(0, pump)
            _emit("ready")
            root.mainloop()
    except Exception as error:
        _emit("error", message=f"{type(error).__name__}: {error}")
        return 1
    finally:
        if box is not None:
            box.destroy()
        if root is not None:
            root.destroy()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
