from pathlib import Path

import pytest

from easy_uiauto.knowledge_library import PAGE_ID
from easy_uiauto.mcp import knowledge
from test_knowledge_library import library,add_control
from test_dynamic_regions import popup


def item(record,kind="control",app_id="demo"):
    return dict(app_id=app_id,kind=kind,id=record["id"],page_id=record.get("page_id",PAGE_ID))


def active_files(store):
    directory=store.directory("demo")
    paths=[*(p for p,_r in knowledge.iter_records(directory)),directory/".easy_uiauto/index.json",directory/"operations/UI-CLI.md"]
    return {p:p.read_bytes() for p in paths if p.is_file()}


def test_many_controls_are_deduplicated_and_index_refreshed_once(library,monkeypatch):
    original=add_control(library);records=[]
    for i in range(12):
        record={**original,"id":f"control-{i}","status":"quarantined" if i%2 else "observed"}
        knowledge.save_control(library.directory("demo"),record);records.append(record)
    refresh=library._refresh;calls=[]
    def counted(directory):calls.append(directory);return refresh(directory)
    monkeypatch.setattr(library,"_refresh",counted)
    paths=library.archive_many([*(item(r) for r in records),item(records[0])])
    assert len(paths)==12 and all(p.is_file() for p in paths)
    assert len({p.parents[1] for p in paths})==1
    assert calls==[library.directory("demo")]
    assert [r["id"] for r in library.read("demo")["controls"]]==[original["id"]]
    assert [r["id"] for r in knowledge.load_index(library.directory("demo"))["controls"]]==[original["id"]]


def test_nonrecursive_archive_keeps_unselected_descendants(library):
    first=add_control(library)
    other={**first,"id":"other","navigation_path":[first["id"]]}
    knowledge.save_control(library.directory("demo"),other)
    identifier=library.save_region("demo","区域",dict(left=0,top=0,right=1,bottom=.5))
    region=next(r for r in library.regions("demo") if r["id"]==identifier)
    library.archive_many([item(region,"region"),item(first)],recursive=False)
    saved=library.read("demo")
    assert not any(r["id"]==identifier for r in saved["regions"])
    assert len(saved["controls"])==1
    assert saved["controls"][0]["id"]=="other" and saved["controls"][0]["region_id"]=="unassigned"
    assert saved["controls"][0]["navigation_path"]==[first["id"]]
    assert saved["controls"][0]["location"]==other["location"]


@pytest.mark.parametrize("invalid",[dict(kind="control",id="missing"),dict(kind="application",id="demo"),
    dict(kind="region",id="classification"),dict(kind="region",id="old-region",page_id="wrong-page")])
def test_entire_batch_is_validated_before_archiving(library,invalid):
    first=add_control(library)
    knowledge.save_region(library.directory("demo"),dict(id="old-region",page_id=PAGE_ID,app_id="demo",source="ai"))
    knowledge.save_region(library.directory("demo"),dict(id="classification",page_id=PAGE_ID,app_id="demo",classification_only=True))
    before=active_files(library)
    with pytest.raises(ValueError):library.archive_many([item(first),dict(app_id="demo",**invalid)])
    assert active_files(library)==before
    assert not list((library.directory("demo")/".easy_uiauto/archive").glob("**/*.md"))


@pytest.mark.parametrize("source",["ai","visual-first","","manual","local-visual-container"])
@pytest.mark.parametrize("assignment",["manual-selection","expanded-popup","automatic"])
def test_archive_real_regions_preserves_other_regions_and_detaches_retained_controls(library,source,assignment):
    first=add_control(library)
    directory=library.directory("demo")
    region=dict(id=PAGE_ID+".old-region",page_id=PAGE_ID,app_id="demo",source=source,name="旧区域")
    other={**region,"id":PAGE_ID+".keep-region", "name":"保留区域"}
    knowledge.save_region(directory,region);knowledge.save_region(directory,other)
    first.update(region_id=region["id"],region_source=assignment,command=PAGE_ID+".old-region.open")
    second={**first,"id":"keep-control","region_id":"keep-region","command":PAGE_ID+".keep-region.open"}
    collision={**first,"id":"unassigned-control","region_id":"unassigned","command":PAGE_ID+".unassigned.open"}
    for control in (first,second,collision):knowledge.save_control(directory,control)
    before={r["id"]:r for r in library.read("demo")["controls"]}
    archived=library.archive("demo","region",region["id"],recursive=False)
    library.sync_dynamic_regions("demo")
    snapshot=library.read("demo");after={r["id"]:r for r in snapshot["controls"]}
    assert archived.is_file() and [r["id"] for r in snapshot["regions"]]==[other["id"]]
    assert after[second["id"]]==before[second["id"]] and after[collision["id"]]==before[collision["id"]]
    retained=after[first["id"]]
    assert retained["region_id"]=="unassigned" and retained["command"].startswith(PAGE_ID+".unassigned.")
    assert len({r["command"] for r in after.values()})==3
    for key in ("location","status","region_source","normalized_rect"):
        assert retained[key]==before[first["id"]][key]
    assert not any(r["id"]==region["id"] for r in knowledge.load_index(directory)["regions"])


@pytest.mark.parametrize("fail_at",["move","refresh"])
def test_write_failure_rolls_back_moves_assignments_and_index(library,monkeypatch,fail_at):
    first=add_control(library)
    second={**first,"id":"second"};knowledge.save_control(library.directory("demo"),second)
    identifier=library.save_region("demo","区域",dict(left=0,top=0,right=1,bottom=.5))
    region=next(r for r in library.regions("demo") if r["id"]==identifier)
    before=active_files(library)
    if fail_at=="move":
        replace=Path.replace;attempts=[]
        def fail_second(path,destination):
            if "archive" in destination.parts:
                attempts.append(path)
                if len(attempts)==2:raise PermissionError("locked record")
            return replace(path,destination)
        monkeypatch.setattr(Path,"replace",fail_second)
    else:
        def denied(_d):raise PermissionError("index locked")
        monkeypatch.setattr(library,"_refresh",denied)
    with pytest.raises(RuntimeError,match="原记录已恢复"):
        library.archive_many([item(region,"region"),item(first)])
    assert active_files(library)==before
    assert not list((library.directory("demo")/".easy_uiauto/archive").glob("**/*.md"))


def test_batch_supports_multiple_software_and_page_states(library,monkeypatch):
    first=add_control(library)
    library.bind("other","第二个软件",dict(title="Other",class_name="Other"))
    second={**first,"id":"second","app_id":"other"};knowledge.save_control(library.directory("other"),second)
    region=dict(id="state-view.work",app_id="demo",page_id="state-view",source="manual",name="绘图",
                normalized_rect=dict(left=0,top=0,right=1,bottom=1))
    knowledge.save_region(library.directory("demo"),region)
    refresh=library._refresh;calls=[]
    monkeypatch.setattr(library,"_refresh",lambda directory:(calls.append(directory),refresh(directory))[-1])
    library.archive_many([item(first),item(second,app_id="other"),item(region,"region")])
    assert not library.read("demo")["controls"] and not library.read("other")["controls"]
    assert calls==[library.directory("demo"),library.directory("other")]


def test_multiple_removed_dynamic_ancestors_are_skipped_for_retained_children(popup):
    store,records,_static=popup
    from easy_uiauto.mcp.dynamic_regions import build_tree
    index=store.read("demo");graph=build_tree(index["controls"],index["regions"])
    root=graph["containers"]["popup"]
    child=store.save_dynamic_region("demo","子",dict(left=.1,top=.2,right=.6,bottom=.35),parent_region_id=root)
    grandchild=store.save_dynamic_region("demo","孙",dict(left=.12,top=.21,right=.58,bottom=.34),parent_region_id=child)
    regions={r["id"]:r for r in store.read("demo")["regions"]}
    store.archive_many([item(regions[root],"region"),item(regions[child],"region")],recursive=False)
    after=store.read("demo");kept=next(r for r in after["regions"] if r["id"]==grandchild)
    assert not kept.get("parent_region_id") and kept["parent_control_id"]=="file"
    assert any(r["id"]=="popup" and r["region_behavior"]=="control" for r in after["controls"])
    assert all(any(r["id"]==old["id"] for r in after["controls"]) for old in records)


@pytest.mark.parametrize("selection",["region","opener"])
def test_recursive_archive_removes_nested_popup_and_hidden_sources_without_resurrection(popup,selection):
    store,controls,static=popup
    from easy_uiauto.mcp.dynamic_regions import build_tree
    directory=store.directory("demo")
    graph=build_tree(store.read("demo")["controls"],store.read("demo")["regions"])
    root=graph["containers"]["popup"]
    child=store.save_dynamic_region("demo","子",dict(left=.1,top=.2,right=.6,bottom=.35),parent_region_id=root)
    grandchild=store.save_dynamic_region("demo","孙",dict(left=.12,top=.21,right=.58,bottom=.34),parent_region_id=child)
    # A same-named toolbar command and a route reference outside the subtree survive.
    sibling={**controls[0],"id":"toolbar-copy","navigation_path":[],"hierarchy_parent":""}
    reference={**controls[2],"id":"referenced-elsewhere","hierarchy_parent":"",
               "entry_steps":[dict(action="click",target="file")]}
    for record in (sibling,reference):knowledge.save_control(directory,record)
    before=store.read("demo");regions={r["id"]:r for r in before["regions"]}
    selected=([item(regions[root],"region"),item(regions[child],"region")] if selection=="region"
              else [item(controls[0])])
    preview=store.expand_archive_items(selected,{"demo":before})
    paths=store.archive_many(selected,expected_items=preview)
    assert len(paths)==len(preview)==(6 if selection=="region" else 7)
    assert len(set(paths))==len(paths) and all(p.is_file() for p in paths)
    store.sync_dynamic_regions("demo")
    after=store.read("demo")
    assert {r["id"] for r in after["regions"]}=={static}
    assert {r["id"] for r in after["controls"]}==({"file","toolbar-copy","referenced-elsewhere"}
                                                 if selection=="region" else {"toolbar-copy","referenced-elsewhere"})
    for identifier in ("toolbar-copy","referenced-elsewhere"):
        assert next(r for r in after["controls"] if r["id"]==identifier)==next(r for r in before["controls"] if r["id"]==identifier)
    assert all(store.asset("demo",r["image"]).is_file() for r in controls)
    assert not {root,child,grandchild}&{r["id"] for r in knowledge.load_index(directory)["regions"]}


def test_recursive_archive_follows_native_and_manual_control_hierarchy(library):
    from easy_uiauto.mcp.protocol import location_from_xpath
    parent=add_control(library);directory=library.directory("demo")
    native={**parent,"id":"native-child","location":location_from_xpath([*parent["location"]["Xpath"],
            dict(ControlType="TextControl",Name="Caption")])}
    manual={**parent,"id":"manual-child","hierarchy_parent":"native-child","location":location_from_xpath([
            *parent["location"]["Xpath"][:-1],dict(ControlType="ButtonControl",Name="Manual")])}
    unrelated={**parent,"id":"sibling","location":location_from_xpath([
            *parent["location"]["Xpath"][:-1],dict(ControlType="ButtonControl",Name="Other")])}
    for record in (native,manual,unrelated):knowledge.save_control(directory,record)
    paths=library.archive_many([item(parent),item(manual)])
    assert len(paths)==3
    assert [r["id"] for r in library.read("demo")["controls"]]==["sibling"]


def test_recursive_archive_includes_owned_popup_page_but_preserves_other_states(library):
    parent=add_control(library);directory=library.directory("demo")
    region=dict(id="popup-page.form",app_id="demo",page_id="popup-page",name="对话框",
                learning_source="deep-learning",window_scope=dict(kind="owned"),parent_control_id=parent["id"])
    knowledge.save_region(directory,region)
    child={**parent,"id":"dialog-field","page_id":"popup-page","region_id":"form"}
    other={**child,"id":"other-state","page_id":"unrelated-page"}
    for record in (child,other):knowledge.save_control(directory,record)
    paths=library.archive_many([item(parent)])
    assert len(paths)==3
    assert not library.read("demo")["regions"]
    assert [r["id"] for r in library.read("demo")["controls"]]==[other["id"]]


def test_changed_subtree_requires_new_confirmation_before_any_write(library):
    parent=add_control(library);directory=library.directory("demo")
    selected=[item(parent)]
    preview=library.expand_archive_items(selected,{"demo":library.read("demo")})
    knowledge.save_control(directory,{**parent,"id":"new-child","hierarchy_parent":parent["id"]})
    before=active_files(library)
    with pytest.raises(ValueError,match="子节点已变化"):
        library.archive_many(selected,expected_items=preview)
    assert active_files(library)==before
    assert not list((directory/".easy_uiauto/archive").glob("**/*.md"))
