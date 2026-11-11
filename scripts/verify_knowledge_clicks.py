"""Verify GUI click and File -> New replay on the disposable native menu fixture."""

from __future__ import annotations

import argparse
import json
import time
import tkinter as tk
from pathlib import Path

from PIL import Image

from easy_uiauto.knowledge_gui import KnowledgeWindow
from easy_uiauto.mcp import native_windows


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--events", type=Path, required=True)
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    vault = output / "vault"
    if (vault / "applications").exists(): raise ValueError("Use a new disposable vault")
    targets = [w for w in native_windows.enumerate_windows(visible_only=True)
               if w["title"] == "Knowledge Menu Replay Target"]
    assert len(targets) == 1, "Start exactly one --menu fixture first"
    target = targets[0]
    baseline = json.loads(args.events.read_text(encoding="utf-8")) if args.events.exists() else []
    root = tk.Tk()
    gui = KnowledgeWindow(root, vault)
    root.title("知识库点击流程验收")
    import easy_uiauto.knowledge_gui as gui_module
    errors = []
    gui_module.messagebox.showerror = lambda title, message, **_kw: errors.append(f"{title}: {message}")
    gui_module.messagebox.showinfo = lambda *_a, **_kw: None
    gui.enrich.set(False)
    root.update()
    app_id = "menu-replay-acceptance"
    library = gui.library
    outcomes = {}

    def await_task(timeout=100):
        assert gui.task is not None, "No worker started: "+repr(errors)
        deadline = time.monotonic()+timeout
        while gui.task and time.monotonic() < deadline:
            root.update(); time.sleep(.025)
        assert gui.task is None, "Worker still running: "+gui.status.get()
        root.update()
        assert not errors, errors

    def select(record):
        root.focus_force()
        root.update()
        gui.tree.selection_set(f"control:{app_id}:{record['id']}")
        gui.select(); root.update()

    def events():
        return json.loads(args.events.read_text(encoding="utf-8"))[len(baseline):]

    try:
        library.bind(app_id, "菜单回放测试软件", target)
        gui.bound(app_id, target); await_task(35)
        automatic = library.layout_regions(app_id)
        assert automatic and all(r.get("layout_origin", "").startswith("auto-") for r in automatic)
        assert all(not r.get("edited_by_user") for r in automatic)
        (output / "automatic-regions.json").write_text(json.dumps(automatic, ensure_ascii=False, indent=2), encoding="utf-8")
        gui.confirm()
        gui.buttons["3  开始学习"].invoke(); await_task()
        records = library.read(app_id)["controls"]
        save = next(r for r in records if r["name"] == "Save" and r["control_type"] == "ButtonControl")
        file = next(r for r in records if r["name"] == "File" and r["control_type"] == "MenuItemControl")
        select(save)
        gui.buttons["点击"].invoke(); await_task(35)
        assert events().count("save") == 1, events()
        outcomes["ordinary_click"] = library.read(app_id)["controls"]
        outcomes["ordinary_click"] = next(r["last_click_execution"] for r in outcomes["ordinary_click"] if r["id"] == save["id"])
        assert outcomes["ordinary_click"]["ok"]
        print("ordinary click passed", flush=True)

        select(file)
        gui.buttons["学习展开控件"].invoke(); await_task()
        records = library.read(app_id)["controls"]
        (output / "after-expand.json").write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
        new = next(r for r in records if r["name"] == "New")
        assert new.get("navigation_path") == [file["id"]], new
        assert new["visual_observation"]["phase"] == "located"
        reference_keys = ("image", "image_sha256", "image_variants", "visual_observation")
        learned_reference = {key: new[key] for key in reference_keys}
        learned_pixels = library.asset(app_id, new["image"]).read_bytes()
        assert events().count("new") == 0, events()
        assert events().count("file-open") == 1, events()
        print("expanded menu learning passed", flush=True)
        select(new)
        graph = gui.region_trees[app_id]
        popup = next(r for r in records if r.get("control_type") == "MenuControl" and r.get("navigation_path") == [file["id"]])
        dynamic_id = graph["containers"][popup["id"]]
        dynamic_key = f"region:{app_id}:{dynamic_id}"
        assert gui.tree.parent(dynamic_key) == f"control:{app_id}:{file['id']}"
        assert gui.tree.parent(f"control:{app_id}:{new['id']}") == dynamic_key
        assert not gui.tree.exists(f"control:{app_id}:{popup['id']}")
        popup_reference = next(r for r in library.read(app_id)["regions"] if r["id"] == dynamic_id)
        popup_pixels = library.asset(app_id, popup_reference["image"]).read_bytes()
        # Default acceptance follows the automatic path; no hand drawn demo groups.
        opened = next(r for r in records if r["name"] == "Open" and r.get("navigation_path") == [file["id"]])
        assert gui.tree.parent(f"control:{app_id}:{opened['id']}") == dynamic_key
        assert len(graph["dynamic"]) == 1
        select(new)
        assert "File" in gui.navigation_summary.cget("text"), gui.navigation_summary.cget("text")
        gui.buttons["点击"].invoke(); await_task(35)
        assert gui.observation_mode.get() == "本次执行结果"
        assert gui.display_after and not gui.viewer.controls
        records = library.read(app_id)["controls"]
        new = next(r for r in records if r["id"] == new["id"])
        execution = new["last_click_execution"]
        assert execution["ok"] and [s["control_id"] for s in execution["steps"]] == [file["id"], new["id"]]
        assert events().count("new") == 1, events()
        assert events().count("file-open") == (1 if execution["steps"][0].get("skipped") else 2), events()
        outcomes["menu_click"] = execution
        before, after = execution["steps"][-1]["before"], execution["steps"][-1]["after"]
        assert before["phase"] == "before-action" and after["phase"] == "after-action"
        assert {key: new[key] for key in reference_keys} == learned_reference
        assert library.asset(app_id, new["image"]).read_bytes() == learned_pixels
        assert before["purpose"] == after["purpose"] == "execution-evidence"
        with Image.open(library.asset(app_id, before["context_image"])) as source:
            frame = source.convert("RGB")
        n = before["normalized_rect"]
        box = tuple(round(n[k]*(frame.width if k in {"left", "right"} else frame.height))
                    for k in ("left", "top", "right", "bottom"))
        with Image.open(library.asset(app_id, before["control_image"])) as source:
            assert frame.crop(box).tobytes() == source.convert("RGB").tobytes(), "Runtime crop must use execution-before pixels"
        frame.save(output / "new-before.png")
        with Image.open(library.asset(app_id, after["context_image"])) as source:
            result_frame = source.convert("RGB")
        assert frame.tobytes() != result_frame.tobytes(), "Closed menu must differ from the pre-click frame"
        result_frame.save(output / "new-after.png")
        gui.observation_mode.set("本次执行结果"); gui.show_observation()
        assert gui.display_after and not gui.viewer.controls
        gui.observation_mode.set("本次执行前"); gui.show_observation()
        assert not gui.display_after and gui.viewer.controls
        gui.observation_mode.set("学习参考图"); gui.show_observation()
        assert gui.display_observation == learned_reference["visual_observation"]
        print("File -> New replay passed", flush=True)

        gui.buttons["3  开始学习"].invoke(); await_task()
        reloaded = next(r for r in library.read(app_id)["controls"] if r["id"] == new["id"])
        assert reloaded.get("navigation_path") == [file["id"]]
        assert reloaded["status"] != "quarantined", reloaded
        assert {key: reloaded[key] for key in reference_keys} == learned_reference
        select(reloaded)
        gui.buttons["点击"].invoke(); await_task(35)
        assert events().count("new") == 2, events()
        second = next(r for r in library.read(app_id)["controls"] if r["id"] == new["id"])["last_click_execution"]
        assert not second["steps"][0].get("skipped"), second
        assert events().count("file-open") == (2 if execution["steps"][0].get("skipped") else 3), events()
        outcomes["menu_closed_click"] = second
        final_record = next(r for r in library.read(app_id)["controls"] if r["id"] == new["id"])
        assert {key: final_record[key] for key in reference_keys} == learned_reference
        assert library.asset(app_id, final_record["image"]).read_bytes() == learned_pixels
        final_regions = {r["id"]: r for r in library.read(app_id)["regions"]}
        assert library.asset(app_id, final_regions[dynamic_id]["image"]).read_bytes() == popup_pixels
        assert len(gui.region_trees[app_id]["dynamic"]) == 1
        assert all(r.get("layout_origin", "").startswith("auto-") and not r.get("edited_by_user")
                   for r in library.regions(app_id))
        result = dict(ok=True, vault=str(vault), ordinary_click_once=True,
                      popup_learned_without_clicking_leaf=True, learned_prerequisite=True,
                      menu_replay_order=True, repeat_scan_retains_path=True,
                      replay_after_rescan=True, already_open_entry_skipped=bool(execution["steps"][0].get("skipped")),
                      learning_reference_preserved=True, runtime_crop_uses_pre_click_pixels=True, after_image_separate=True,
                      after_view_has_no_target_box=True, runtime_result_shown_automatically=True,
                      rescan_retains_learning_frame=True,
                      menu_item_nested_under_entry=True, popup_is_dynamic_region=True,
                      automatic_initial_layout=True, no_manual_regions_created=True,
                      no_per_command_regions=True, region_does_not_add_click=True,
                      dynamic_region_reference_preserved=True,
                      api_requests=0, events=events(), executions=outcomes)
        (output / "acceptance.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
    finally:
        if not gui.closed:
            if gui.task: gui.task.close(); gui.task = None
            gui.dirty = False; gui.close()


if __name__ == "__main__": main()
