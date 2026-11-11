from copy import deepcopy

from PIL import Image, ImageDraw

from easy_uiauto.mcp import automatic_layout as layout, manual_layout
from test_knowledge_library import library  # Shared disposable vault fixture.


def node(kind, name, bounds, children=()):
    return dict(ControlType=kind, Name=name, bounds=dict(zip(layout.KEYS, bounds)), children=list(children))


def test_uia_bars_and_panels_form_regions_without_one_region_per_command():
    image = Image.new("RGB", (800, 600), "white")
    children = [node("TitleBarControl", "Demo", (0, 0, 800, 30)),
        node("MenuBarControl", "", (0, 30, 800, 55)),
        node("ToolBarControl", "", (0, 55, 800, 95), [node("ButtonControl", name, (i*50, 60, i*50+40, 90))
            for i, name in enumerate(("New", "Open", "Save"))]),
        node("TreeControl", "项目", (0, 95, 180, 575)),
        node("DocumentControl", "编辑器", (180, 95, 800, 575)),
        node("StatusBarControl", "", (0, 575, 800, 600))]
    result = layout.propose(image, node("WindowControl", "Demo", (0, 0, 800, 600), children))
    names = {r["name"] for r in result["regions"]}
    assert {"标题栏", "菜单栏", "工具栏", "项目", "编辑器", "状态栏"} <= names
    assert not {"New", "Open", "Save"} & names
    assert len(result["regions"]) == 6
    manual_layout.validate_regions(result["regions"])


def test_duplicate_containers_and_whole_window_wrappers_do_not_create_extra_regions():
    image = Image.new("RGB", (800, 600))
    panel = node("DocumentControl", "Editor", (0, 80, 800, 600))
    tree = node("WindowControl", "Demo", (0, 0, 800, 600), [
        node("PaneControl", "Wrapper", (0, 0, 800, 600), [panel]), panel,
        node("ButtonControl", "New", (10, 40, 50, 65))])
    result = layout.propose(image, tree)
    assert sum(r["name"] == "Editor" for r in result["regions"]) == 1
    assert not any(r["name"] in {"Wrapper", "New"} for r in result["regions"])


def test_textbox_borders_and_blank_spacing_do_not_fragment_an_identified_layout():
    image = Image.new("RGB", (800, 600), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((15, 220, 780, 440), outline="black")
    draw.rectangle((15, 500, 780, 530), fill="#eeeeee")
    tree = node("WindowControl", "Demo", (0, 0, 800, 600), [
        node("TitleBarControl", "Demo", (16, 0, 800, 30)),
        node("MenuBarControl", "", (0, 30, 800, 55)),
        *(node("ButtonControl", name, (20+i*130, 95, 130+i*130, 135)) for i, name in enumerate(("Open", "Save", "Options"))),
        node("EditControl", "", (15, 220, 780, 440))])
    result = layout.propose(image, tree)
    assert {r["name"] for r in result["regions"]} == {"标题栏", "菜单栏", "工具栏", "内容区"}
    assert len(result["regions"]) == 4
    toolbar = next(r for r in result["regions"] if r["name"] == "工具栏")
    assert toolbar["normalized_rect"]["top"] == 55/600


def test_pixel_boundaries_produce_coarse_nonoverlapping_proposal_and_blank_is_one_region():
    image = Image.new("RGB", (800, 600), "white")
    assert len(layout.propose(image)["regions"]) == 1
    draw = ImageDraw.Draw(image)
    draw.line((0, 100, 799, 100), fill="black", width=1)
    draw.line((200, 101, 200, 599), fill="black", width=1)
    result = layout.propose(image)
    assert 2 <= len(result["regions"]) <= 6
    assert result["method"] == "pixel-dividers"
    manual_layout.validate_regions(result["regions"])


def test_screen_origin_is_removed_and_repeated_proposals_are_stable():
    image = Image.new("RGB", (400, 300), "white")
    tree = node("WindowControl", "Demo", (-500, 100, -100, 400), [
        node("TitleBarControl", "Demo", (-500, 100, -100, 130))])
    first = layout.propose(image, tree, origin=(-500, 100))
    assert first == layout.propose(image, tree, origin=(-500, 100))
    title = next(r for r in first["regions"] if r["name"] == "标题栏")
    assert title["normalized_rect"] == dict(left=0, top=0, right=1, bottom=.1)


def test_repartition_preserves_user_adjustments_and_replaces_only_untouched_suggestions(library):
    image = Image.new("RGB", (400, 300), "white")
    proposal = layout.propose(image)
    result = library.apply_automatic_layout("demo", proposal)
    assert result["suggested"] == 1 and result["preserved"] == 0
    initial = library.layout_regions("demo")[0]
    library.save_region("demo", "人工调整", dict(left=0, top=0, right=.3, bottom=1), initial["id"])
    edited = next(r for r in library.layout_regions("demo") if r["name"] == "人工调整")
    original = deepcopy(edited)
    result = library.apply_automatic_layout("demo", proposal)
    assert result["preserved"] == 1 and result["suggested"] == 1
    saved = library.layout_regions("demo")
    assert next(r for r in saved if r["id"] == original["id"]) == original
    manual_layout.validate_regions(saved)
    assert not library.page("demo")["layout_confirmed"]
    library.confirm_layout("demo")
    assert library.learning_layout("demo", dict(handle=1, process_id=2, class_name="Window"))["regions"]
    library.apply_automatic_layout("demo", proposal)
    assert len(library.layout_regions("demo")) == 2


def test_auto_region_can_be_renamed_through_the_gui_and_is_then_preserved(library):
    import tkinter as tk
    from easy_uiauto.knowledge_gui import KnowledgeWindow

    library.apply_automatic_layout("demo", layout.propose(Image.new("RGB", (400, 300), "white")))
    region = library.regions("demo")[0]
    root = tk.Tk(); root.withdraw()
    gui = None
    try:
        gui = KnowledgeWindow(root, library.root)
        gui.tree.selection_set("region:demo:"+region["id"]); gui.select()
        gui.name.set("我的工作区")
        assert gui.save()
        assert library.regions("demo")[0]["edited_by_user"]
        assert library.regions("demo")[0]["name"] == "我的工作区"
        assert "添加区域" in gui.buttons and "修改区域" in gui.buttons
        assert "自动划分区域" not in gui.buttons
    finally:
        if gui: gui.dirty = False; gui.close()
        else: root.destroy()
