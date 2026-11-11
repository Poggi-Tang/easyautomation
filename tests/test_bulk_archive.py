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


def test_selected_parent_and_child_do_not_delete_unselected_descendants(library):
    first=add_control(library)
    other={**first,"id":"other","navigation_path":[first["id"]]}
    knowledge.save_control(library.directory("demo"),other)
    identifier=library.save_region("demo","区域",dict(left=0,top=0,right=1,bottom=.5))
    region=next(r for r in library.regions("demo") if r["id"]==identifier)
    library.archive_many([item(region,"region"),item(first)])
    saved=library.read("demo")
    assert not any(r["id"]==identifier for r in saved["regions"])
    assert len(saved["controls"])==1
    assert saved["controls"][0]["id"]=="other" and saved["controls"][0]["region_id"]=="unassigned"
    assert saved["controls"][0]["navigation_path"]==[first["id"]]
    assert saved["controls"][0]["location"]==other["location"]


@pytest.mark.parametrize("invalid",[dict(kind="control",id="missing"),dict(kind="application",id="demo"),
    dict(kind="region",id="old-region")])
def test_entire_batch_is_validated_before_archiving(library,invalid):
    first=add_control(library)
    knowledge.save_region(library.directory("demo"),dict(id="old-region",page_id=PAGE_ID,app_id="demo",source="ai"))
    before=active_files(library)
    with pytest.raises(ValueError):library.archive_many([item(first),dict(app_id="demo",**invalid)])
    assert active_files(library)==before
    assert not list((library.directory("demo")/".easy_uiauto/archive").glob("**/*.md"))


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
    store.archive_many([item(regions[root],"region"),item(regions[child],"region")])
    after=store.read("demo");kept=next(r for r in after["regions"] if r["id"]==grandchild)
    assert not kept.get("parent_region_id") and kept["parent_control_id"]=="file"
    assert any(r["id"]=="popup" and r["region_behavior"]=="control" for r in after["controls"])
    assert all(any(r["id"]==old["id"] for r in after["controls"]) for old in records)
