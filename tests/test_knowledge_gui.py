from types import SimpleNamespace

from PIL import Image

from easy_uiauto.knowledge_gui import KnowledgeWindow, RegionCanvas, WindowChooser
from easy_uiauto.knowledge_library import KnowledgeLibrary, PAGE_ID
from easy_uiauto.mcp import knowledge
from test_knowledge_library import add_control


def test_canvas_maps_zoom_and_scroll_to_original_image_coordinates():
    # Exercise mapping without capturing or clicking the user's desktop.
    viewer = object.__new__(RegionCanvas)
    viewer.image = Image.new("RGB", (800, 600))
    viewer.scale = .5
    viewer.canvas = SimpleNamespace(canvasx=lambda x: x+100, canvasy=lambda y: y+50)
    assert viewer.point(100, 100) == (400, 300)
    assert viewer.point(-500, 9999) == (0, 600)


def test_gui_tree_keeps_single_controls_directly_under_region(tmp_path):
    import tkinter as tk

    library = KnowledgeLibrary(tmp_path)
    library.bind("demo", "测试软件", dict(title="Test", class_name="Window"))
    source = tmp_path / "source.png"; Image.new("RGB", (400, 300)).save(source)
    library.save_capture("demo", dict(image=str(source), width=400, height=300, window_name="Test", window_class="Window"))
    add_control(library, status="verified")
    identifier = library.save_region("demo", "工具栏", dict(left=0, top=0, right=1, bottom=.5))
    root = tk.Tk(); root.withdraw()
    window = None
    try:
        window = KnowledgeWindow(root, tmp_path)
        assert window.enrich.get() is False
        assert window.tree.get_children("") == ("vault",)
        app = window.tree.get_children("vault")
        assert app == ("app:demo",)
        assert window.tree.get_children(app[0]) == ("usage-root:demo", "controls-root:demo", "actions-root:demo")
        regions = window.tree.get_children("controls-root:demo")
        assert regions == (f"region:demo:{identifier}",)
        controls = window.tree.get_children(regions[0])
        assert controls == ("control:demo:control-one",)
        assert not window.tree.get_children(controls[0])
        window.tree.selection_set(controls[0]); window.select()
        assert '"WindowName"' in window.location.get("1.0", "end")
        assert "find_control" in window.code.get("1.0", "end")
        assert window.selected[2]["name"] == "Open"
        window.name.set("业务按钮")
        assert window.save()
        updated = library.read("demo")["controls"][0]
        assert updated["name"] == "Open" and updated["semantic_name"] == "业务按钮"
        assert updated["status"] == "verified"
    finally:
        if window:
            window.dirty = False
            window.close()
        else:
            root.destroy()


def test_nested_tree_manual_reparenting_preserves_locator_and_click_path(tmp_path):
    import tkinter as tk
    from easy_uiauto.mcp.protocol import location_from_xpath

    library = KnowledgeLibrary(tmp_path)
    library.bind("demo", "Demo", dict(title="Test", class_name="Window"))
    source = tmp_path / "source.png"; Image.new("RGB", (400, 300)).save(source)
    library.save_capture("demo", dict(image=str(source), width=400, height=300, window_name="Test", window_class="Window"))
    file = add_control(library)
    region = library.save_region("demo", "菜单栏", dict(left=0, top=0, right=1, bottom=.5))
    file = library.read("demo")["controls"][0]
    def menu(identifier, path):
        return {**file, "id": identifier, "name": identifier, "semantic_name": identifier,
                "navigation_path": path, "location": location_from_xpath([
                    dict(ControlType="WindowControl", Name="Test"), dict(ControlType="MenuItemControl", Name=identifier)])}
    recent, leaf = menu("recent", [file["id"]]), menu("project", [file["id"], "recent"])
    leaf["region_id"] = "unassigned"
    knowledge.save_control(library.directory("demo"), recent)
    knowledge.save_control(library.directory("demo"), leaf)
    root = tk.Tk(); root.withdraw()
    gui = None
    try:
        gui = KnowledgeWindow(root, tmp_path)
        file_key, recent_key, leaf_key = ["control:demo:"+identifier for identifier in (file["id"], "recent", "project")]
        assert gui.tree.parent(file_key) == "region:demo:"+region
        assert gui.tree.parent(recent_key) == file_key
        assert gui.tree.parent(leaf_key) == recent_key
        assert gui.hierarchies["demo"]["regions"]["project"] == region
        gui.tree.selection_set(leaf_key); gui.select()
        gui.parent_selection.set("区域直属")
        assert gui.save()
        assert gui.tree.parent(leaf_key) == "region:demo:manual-main.unassigned"
        label = next(label for label, identifier in gui.parent_options.items() if identifier == file["id"])
        gui.parent_selection.set(label)
        assert gui.save()
        assert gui.tree.parent(leaf_key) == file_key
        saved = next(r for r in library.read("demo")["controls"] if r["id"] == "project")
        assert saved["location"] == leaf["location"]
        assert saved["navigation_path"] == leaf["navigation_path"]
        assert saved["region_id"] == "unassigned"
        import pytest
        with pytest.raises(ValueError, match="循环"):
            library.update_control("demo", file["id"], name="文件", description="", aliases=[], hierarchy_parent="project")
    finally:
        if gui:
            gui.dirty = False; gui.close()
        else: root.destroy()


def test_window_selection_updates_generated_name_but_preserves_user_text(tmp_path, monkeypatch):
    import tkinter as tk
    from easy_uiauto.mcp import native_windows

    windows = [dict(handle=i, title=f"Window {i}", class_name="Native", bounds=dict(width=400)) for i in (1, 2)]
    monkeypatch.setattr(native_windows, "enumerate_windows", lambda **_kw: windows)
    root = tk.Tk(); root.withdraw()
    chooser = None
    try:
        chooser = WindowChooser(root, KnowledgeLibrary(tmp_path), "", lambda *_a: None)
        chooser.tree.selection_set("1"); chooser.changed()
        first_id = chooser.app_id.get()
        chooser.tree.selection_set("2"); chooser.changed()
        assert chooser.name.get() == "Window 2" and chooser.app_id.get() != first_id
        chooser.name.set("自定义名称"); chooser.app_id.set("custom")
        chooser.tree.selection_set("1"); chooser.changed()
        assert chooser.name.get() == "自定义名称" and chooser.app_id.get() == "custom"
    finally:
        if chooser: chooser.destroy()
        root.destroy()


def test_learning_and_runtime_views_are_independent_and_click_displays_result(tmp_path, monkeypatch):
    import tkinter as tk
    from easy_uiauto.mcp import image_storage

    library = KnowledgeLibrary(tmp_path)
    library.bind("demo", "Test", dict(title="Test", class_name="Window"))
    source = tmp_path / "layout.png"; Image.new("RGB", (400, 300), "white").save(source)
    library.save_capture("demo", dict(image=str(source), width=400, height=300, window_name="Test", window_class="Window"))
    record = add_control(library)
    directory = library.directory("demo")
    before = image_storage.save_interaction_image(directory, Image.new("RGB", (400, 300), "blue"))
    runtime_before = image_storage.save_interaction_image(directory, Image.new("RGB", (400, 300), "green"))
    after = image_storage.save_interaction_image(directory, Image.new("RGB", (400, 300), "yellow"))
    record["visual_observation"] = {"phase": "before-action", "context_image": before,
        "normalized_rect": dict(left=.2, top=.2, right=.4, bottom=.4),
        "main_rect": dict(left=0, top=0, right=1, bottom=1)}
    record["last_click_execution"] = {"steps": [{"control_id": record["id"],
        "before": {"phase": "before-action", "context_image": runtime_before,
                   "normalized_rect": dict(left=.3, top=.3, right=.5, bottom=.5),
                   "main_rect": dict(left=0, top=0, right=1, bottom=1)},
        "after": {"phase": "after-action", "context_image": after,
                  "main_rect": dict(left=0, top=0, right=1, bottom=1)}}]}
    knowledge.save_control(directory, record)
    root = tk.Tk(); root.withdraw()
    gui = None
    try:
        gui = KnowledgeWindow(root, tmp_path)
        gui.tree.selection_set("control:demo:control-one"); gui.select()
        assert gui.viewer.image.getpixel((0, 0)) == (0, 0, 255)
        assert gui.viewer.controls[0]["normalized_rect"]["left"] == .2
        gui.observation_mode.set("本次执行前"); gui.show_observation()
        assert gui.viewer.image.getpixel((0, 0)) == (0, 128, 0)
        assert gui.viewer.controls[0]["normalized_rect"]["left"] == .3
        gui.observation_mode.set("本次执行结果"); gui.show_observation()
        assert gui.viewer.image.getpixel((0, 0)) == (255, 255, 0)
        assert not gui.viewer.controls and gui.display_after
        gui.observation_mode.set("学习参考图"); gui.show_observation()
        assert gui.viewer.controls and not gui.display_after
        assert gui.viewer.image.getpixel((0, 0)) == (0, 0, 255)
        monkeypatch.setattr(root, "deiconify", lambda: None)
        monkeypatch.setattr(root, "lift", lambda: None)
        gui.task = SimpleNamespace(poll=lambda: {"state": "completed", "worker_running": False,
                                                "result": {"message": "点击已执行"}}, close=lambda: None)
        gui.task_operation, gui.task_app, gui.task_control = "click", "demo", record["id"]
        gui.poll()
        assert gui.observation_mode.get() == "本次执行结果"
        assert gui.viewer.image.getpixel((0, 0)) == (255, 255, 0)
        assert not gui.viewer.controls
        gui.selected[2]["last_click_execution"]["steps"][0].pop("after")
        gui.show_observation()
        assert gui.viewer.image is None
        assert gui.observation_mode.get() == "本次执行结果"  # Never relabel learning as runtime.
    finally:
        if gui:
            gui.dirty = False; gui.close()
        else: root.destroy()
