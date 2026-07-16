from __future__ import annotations

import json
import queue
import sys
import threading
from typing import Any


def _emit(event: str, **data: Any) -> None:
    print(json.dumps({"event": event, **data}, ensure_ascii=False), flush=True)


def main() -> None:
    import tkinter as tk

    from easy_uiauto.draw import ScreenLineBox
    from easy_uiauto.utils import ensure_uiautomation_thread, find_control_by_xpath

    commands: queue.Queue[dict[str, Any]] = queue.Queue()
    root = None
    box = None
    try:
        first_line = sys.stdin.readline()
        if not first_line:
            raise RuntimeError("highlight worker received no initialization command")
        initial = json.loads(first_line)
        if initial.get("cmd") != "init":
            raise ValueError("first highlight worker command must be init")

        ensure_uiautomation_thread()
        root = tk.Tk()
        root.withdraw()
        style = initial.get("style") or {}
        box = ScreenLineBox(
            root=root,
            control=None,
            mode="track",
            interval_ms=50,
            color=style.get("color", "#FF0000"),
            line_width=style.get("line_width", 2),
            alpha=style.get("alpha", 1.0),
            topmost=True,
        )
        if initial.get("rect"):
            box.set_rect(initial["rect"])

        def read_commands() -> None:
            try:
                for line in sys.stdin:
                    commands.put(json.loads(line))
            except BaseException as exc:
                commands.put({"cmd": "reader_error", "error": str(exc)})
            finally:
                commands.put({"cmd": "stop"})

        threading.Thread(target=read_commands, name="highlight-command-reader", daemon=True).start()

        def pump() -> None:
            while True:
                try:
                    command = commands.get_nowait()
                except queue.Empty:
                    break
                name = command.get("cmd")
                try:
                    if name == "rect":
                        box.set_rect(command.get("rect"))
                    elif name == "set_xpath":
                        control = find_control_by_xpath(
                            command.get("xpath"), use_cache=False
                        )
                        box.set_control(control)
                    elif name == "style":
                        box.set_style(**(command.get("style") or {}))
                    elif name == "reader_error":
                        _emit("error", message=command.get("error"))
                    elif name == "stop":
                        box.destroy()
                        root.destroy()
                        return
                except Exception as exc:
                    _emit("error", message=f"highlight command failed: {exc}")
            root.after(16, pump)

        pump()
        _emit("ready")
        root.mainloop()
    except BaseException as exc:
        _emit("error", message=f"{type(exc).__name__}: {exc}")
        raise SystemExit(1) from exc
    finally:
        if box is not None:
            try:
                box.destroy()
            except Exception:
                pass
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass


if __name__ == "__main__":
    main()
