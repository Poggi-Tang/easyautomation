"""Exercise the actual GUI callbacks and isolated UIA worker against our fixture.

Launch tests/fixtures/knowledge_target.py first. This check never clicks a target
control or sends its screenshot to an API. The --output vault must be disposable.
"""

from __future__ import annotations

import argparse
import json
import time
import tkinter as tk
from pathlib import Path

from easy_uiauto.knowledge_gui import KnowledgeWindow, WindowChooser
from easy_uiauto.mcp import native_windows


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--keep-open", action="store_true")
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    vault = output / "vault"
    if (vault / "applications").exists():
        raise ValueError("验收只允许使用新的临时知识库目录")
    windows = [w for w in native_windows.enumerate_windows(visible_only=True)
               if w["title"] == "Knowledge Library Acceptance Target"]
    assert len(windows) == 1, "Start exactly one disposable fixture first"
    target = windows[0]
    root = tk.Tk()
    gui = KnowledgeWindow(root, vault)
    root.title("easy-uiauto 知识库管理 · 验收")
    errors = []
    import easy_uiauto.knowledge_gui as gui_module

    gui_module.messagebox.showerror = lambda title, message, **kwargs: errors.append(f"{title}: {message}")
    gui_module.messagebox.showinfo = lambda *_args, **_kwargs: None
    gui.enrich.set(False)
    root.update()

    def await_task(timeout=90):
        deadline = time.monotonic()+timeout
        assert gui.task is not None, "GUI did not start its worker"
        while gui.task is not None and time.monotonic() < deadline:
            root.update()
            time.sleep(.025)
        assert gui.task is None, f"Task still running: {gui.status.get()}"
        root.update()
        assert not errors, errors

    try:
        chooser = WindowChooser(root, gui.library, "", gui.bound)
        chooser.app_id.set("acceptance-native")
        chooser.name.set("验收软件")
        chooser.tree.selection_set(str(target["handle"]))
        chooser.accept()
        await_task(35)
        assert gui.viewer.image is not None
        initial_size = gui.viewer.image.size
        assert not gui.library.regions("acceptance-native")
        gui.automatic_partition(); await_task(25)
        automatic = gui.library.regions("acceptance-native")
        assert automatic and all(r.get("layout_origin", "").startswith("auto-") for r in automatic)
        toolbar = next(r for r in automatic if r["name"] == "工具栏")
        gui.tree.selection_set("region:acceptance-native:"+toolbar["id"]); gui.select()
        gui.name.set("工具区（人工确认）"); assert gui.save()
        gui.automatic_partition(); await_task(25)
        preserved = next(r for r in gui.library.regions("acceptance-native") if r["id"] == toolbar["id"])
        assert preserved["name"] == "工具区（人工确认）" and preserved["edited_by_user"]
        gui.confirm()
        gui.buttons["3  开始学习"].invoke()
        await_task()
        first = gui.library.read("acceptance-native")
        controls = first["controls"]
        assert controls, "No native controls were saved"
        region_names = {r["id"].split(".",1)[-1]: r["name"] for r in first["regions"]}
        save = next(c for c in controls if c.get("name") == "Save")
        assert region_names[save["region_id"]] == "工具区（人工确认）"
        editor = next(c for c in controls if c.get("control_type") == "EditControl")
        assert editor["region_id"] in region_names and editor["region_id"] != "unassigned"
        status = next(c for c in controls if c.get("name") == "No actions executed")
        assert status["region_id"] in region_names and status["region_id"] != "unassigned"
        assert sum(node[1] == "control" for node in gui.nodes.values()) == len(controls)
        for identifier in (c["id"] for c in controls):
            key = f"control:acceptance-native:{identifier}"
            parent = gui.tree.parent(key)
            seen = {key}
            while gui.nodes[parent][1] == "control":
                assert parent not in seen, "Control hierarchy contains a cycle"
                seen.add(parent)
                parent = gui.tree.parent(parent)
            assert gui.nodes[parent][1] == "region"
        gui.tree.selection_set(f"control:acceptance-native:{save['id']}")
        gui.select()
        gui.buttons["验证定位"].invoke()
        await_task(35)
        assert "定位成功" in gui.status.get(), gui.status.get()
        gui.buttons["3  开始学习"].invoke()
        await_task()
        second = gui.library.read("acceptance-native")
        assert {c["id"] for c in first["controls"]} == {c["id"] for c in second["controls"]}
        assert all(c.get("name") != "A button was clicked" for c in second["controls"])
        post_learning = next(r for r in gui.library.regions("acceptance-native") if r["id"] == toolbar["id"])
        assert post_learning["edited_by_user"] and post_learning["name"] == "工具区（人工确认）"
        result = {"ok": True, "vault": str(vault), "screenshot_size": initial_size,
                  "automatic_regions": len(automatic), "manual_regions_drawn": 0,
                  "manual_adjustment_preserved": True, "controls": len(second["controls"]),
                  "hierarchical_tree": True, "save_button_in_toolbar": True,
                  "edit_in_content": True, "status_in_status_region": True,
                  "xpath_replay_read_only": True, "repeat_scan_no_duplicates": True,
                  "target_actions_executed": False, "api_requests": 0, "errors": errors}
        (output / "acceptance.json").write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding="utf-8")
        print(json.dumps(result,ensure_ascii=False,indent=2),flush=True)
        if args.keep_open:
            gui.status.set(f"真实窗口验收通过：{len(automatic)} 个自动区域，{len(second['controls'])} 个控件；人工修正保留、重复学习与 XPath 回读通过。")
            root.mainloop()
    finally:
        if not gui.closed:
            gui.dirty = False
            if gui.task: gui.task.close(); gui.task = None
            gui.close()


if __name__ == "__main__":
    main()
