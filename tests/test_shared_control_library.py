"""The legacy MCP path, Python library and GUI must share canonical records."""
import json
import subprocess
import sys

import pytest

from easy_uiauto.knowledge_library import KnowledgeLibrary, PAGE_ID
from easy_uiauto.mcp import knowledge, server
from easy_uiauto.mcp.protocol import location_from_xpath
from test_knowledge_library import library, add_control


def legacy(name="Open"):
    return dict(alias="Demo."+name, app_name="Demo", window_name="Test", notes="Imported observation",
                LOCATION=location_from_xpath([dict(ControlType="WindowControl", Name="Test"),
                    dict(ControlType="ButtonControl", Name=name, foundIndex=1)]))


def test_python_mcp_gui_read_and_search_the_same_record(library, monkeypatch, tmp_path):
    monkeypatch.setenv("EASY_UIAUTO_KNOWLEDGE_DIR", str(tmp_path/"other"))
    monkeypatch.delenv("EASY_UIAUTO_CONTROL_VECTOR_DB_DIR", raising=False)
    captured = legacy()
    first = json.loads(server.upsert_control_record(json.dumps(captured), app_id="demo", vault=str(library.root)))
    assert first["ok"]
    second = library.upsert_control_record(captured, app_id="demo")
    assert second["id"] == first["id"] and second["action"] == "updated"
    gui_records = library.read("demo")["controls"]
    mcp_records = json.loads(server.get_ui_library("demo", vault=str(library.root), section="controls"))["records"]
    search = json.loads(server.search_ui_knowledge("demo", query="Open", vault=str(library.root)))
    assert len(gui_records) == len(mcp_records) == len(search) == 1
    assert gui_records[0]["id"] == mcp_records[0]["id"] == search[0]["id"] == first["id"]
    assert gui_records[0]["status"] == "observed"
    assert gui_records[0]["verification"] == dict(location="pending", image="pending")
    assert "image" not in gui_records[0] and "rect" not in gui_records[0]
    assert json.loads(server.list_ui_commands("demo", vault=str(library.root))) == []


def test_import_preserves_manual_names_paths_images_and_quarantine(library):
    existing = add_control(library, status="quarantined", quarantine=dict(active=True, reason="stale"))
    existing.update(semantic_source="manual", description="User description", image="images/reference.png",
                    entry_steps=[], navigation_path=[], verification=dict(runtime="failed"))
    knowledge.save_control(library.directory("demo"), existing)
    imported = library.upsert_control_record({**legacy(), "status":"verified", "verification":{"image":"passed"}}, app_id="demo")["record"]
    for key in ("semantic_name", "description", "image", "entry_steps", "navigation_path", "quarantine"):
        assert imported[key] == existing[key]
    assert imported["id"] == existing["id"] and imported["status"] == "quarantined"
    assert imported["verification"]["runtime"] == "failed"


def test_ambiguous_state_requires_scope_and_reuses_scoped_record(library):
    first = add_control(library)
    knowledge.save_page(library.directory("demo"), dict(id="second", app_id="demo", name="Tab 2"))
    knowledge.save_control(library.directory("demo"), {**first, "id":"second-control", "page_id":"second"})
    with pytest.raises(ValueError, match="多个状态"):
        library.upsert_control_record(legacy(), app_id="demo")
    result = library.upsert_control_record(legacy(), app_id="demo", page_id="second")
    assert result["id"] == "second-control"
    with pytest.raises(ValueError, match="多个页面"):
        library.upsert_control_record(legacy("New"), app_id="demo")


def test_unique_window_infers_existing_gui_app_and_no_duplicate(library):
    result = library.upsert_control_record(legacy())
    assert result["app_id"] == "demo"
    assert len(library.apps()) == 1
    library.bind("other", "Other", dict(title="Test", class_name="Window"))
    with pytest.raises(ValueError, match="多个软件"):
        library.upsert_control_record(legacy())


def test_json_migration_preview_apply_retry_and_invalid_rows(library, tmp_path):
    path = tmp_path/"old.json"
    path.write_text(json.dumps({"records":[legacy(), {"alias":"missing locator"}, legacy("Save")]}))
    original = path.read_bytes()
    preview = library.import_control_records(path, app_id="demo")
    assert preview["dry_run"] and preview["succeeded"] == 2 and not preview["ok"]
    assert library.read("demo")["controls"] == []
    for _ in range(2):
        result = json.loads(server.import_control_records(str(path), app_id="demo", vault=str(library.root), dry_run=False))
        assert result["succeeded"] == 2 and result["results"][1]["row"] == 2
    assert len(library.read("demo")["controls"]) == 2 and path.read_bytes() == original


def test_index_rebuilds_after_python_record_write(library):
    knowledge.rebuild_index(library.directory("demo"))
    record = add_control(library)
    results = json.loads(server.search_ui_knowledge("demo", vault=str(library.root)))
    assert [r["id"] for r in results] == [record["id"]]
    library.archive("demo", "control", record["id"])
    assert json.loads(server.search_ui_knowledge("demo", vault=str(library.root))) == []


def test_mixed_reports_shared_save_and_failure_without_repeating_action(library, monkeypatch):
    monkeypatch.setattr(server, "CURRENT_MODE", "mixed")
    monkeypatch.setattr(server, "LEARNING_CONTEXT", dict(vault=str(library.root), app_id="demo", page_id=PAGE_ID))
    monkeypatch.setattr(server, "_learning_log", lambda _:None)
    monkeypatch.setattr(server, "_find_control", lambda *a, **k:object())
    monkeypatch.setattr(server, "_record_from_control", lambda *a, **k:legacy())
    calls = []
    @server._action_result()
    def action():
        calls.append(True)
        server._auto_learn_location("Click", legacy()["LOCATION"])
        return "Clicked"
    first = json.loads(action())
    assert first["ok"] and first["learning"][0]["ok"]
    assert library.read("demo")["controls"][0]["id"] == first["learning"][0]["id"]
    def fail(*a, **k):raise OSError("disk unavailable")
    monkeypatch.setattr(KnowledgeLibrary, "upsert_control_record", fail)
    second = json.loads(action())
    assert second["ok"] and not second["learning"][0]["ok"]
    assert "disk unavailable" in second["learning"][0]["error"] and len(calls) == 2


def test_explicit_region_is_shared_with_gui_and_missing_region_rejected(library):
    region = library.save_region("demo", "Toolbar", dict(left=0, top=0, right=1, bottom=.3))
    result = library.upsert_control_record(legacy(), app_id="demo", region_id=region)
    assert result["record"]["region_id"] == region and result["record"]["page_id"] == PAGE_ID
    with pytest.raises(ValueError, match="区域不存在"):
        library.upsert_control_record(legacy("Bad"), app_id="demo", region_id="unknown")


def test_separate_python_process_updates_shared_search_and_revision(library):
    from easy_uiauto.vault_changes import revisions
    before = revisions(library.root)
    script = """import json, sys
from easy_uiauto import KnowledgeLibrary
store = KnowledgeLibrary(sys.argv[1])
result = store.upsert_control_record(json.loads(sys.argv[2]), app_id='demo')
print(result['id'])
"""
    outcome = subprocess.run([sys.executable, "-c", script, str(library.root), json.dumps(legacy())],
                             capture_output=True, text=True, timeout=20, check=True)
    assert revisions(library.root) != before
    result = json.loads(server.search_ui_knowledge("demo", vault=str(library.root)))
    assert result[0]["id"] == outcome.stdout.strip()


def test_jsonl_new_app_roundtrip_and_mode_scope_validation(tmp_path, monkeypatch):
    library = KnowledgeLibrary(tmp_path/"vault")
    path=tmp_path/"old.jsonl";path.write_text(json.dumps(legacy())+"\n"+json.dumps(legacy("Save")))
    result=library.import_control_records(path, dry_run=False)
    assert result["ok"] and len(library.apps())==1
    again=library.import_control_records(path, dry_run=False)
    assert again["ok"] and len(library.read("demo")["controls"])==2
    monkeypatch.setattr(server,"_ensure_learning_console",lambda:None)
    monkeypatch.setattr(server,"_learning_log",lambda _:None)
    monkeypatch.setattr(server,"CURRENT_MODE","operate")
    monkeypatch.setattr(server,"LEARNING_CONTEXT",{})
    bad=json.loads(server.set_mode("mixed",vault=str(library.root),app_id="demo",page_id="missing"))
    assert not bad["ok"] and server.CURRENT_MODE=="operate"
    good=json.loads(server.set_mode("mixed",vault=str(library.root),app_id="demo"))
    assert good["ok"] and json.loads(server.get_mode())["learning_context"]["vault"]==str(library.root)


def test_gui_pick_and_legacy_capture_share_id_without_changing_explicit_mode(library, tmp_path):
    from test_region_picker import picked
    result=picked(tmp_path)
    source={"LOCATION":result["selection"]["location"],"window_name":"Test"}
    imported=library.upsert_control_record(source,app_id="demo")["record"]
    picked_record=library.save_captured_control("demo",result,name="User chosen",region_id="")
    assert picked_record["id"]==imported["id"]
    native=location_from_xpath([{k:v for k,v in step.items() if k!="match_mode"}
                               for step in result["selection"]["location"]["Xpath"]])
    repeated=library.upsert_control_record({"LOCATION":native},app_id="demo")["record"]
    assert repeated["id"]==picked_record["id"] and repeated["semantic_name"]=="User chosen"
    assert repeated["location"]["PARAMETERS"]["match_mode"]=="position"
    assert len(library.read("demo")["controls"])==1


def test_scan_reuses_imported_control_without_automation_id(tmp_path, monkeypatch):
    from easy_uiauto.mcp import scanner
    from test_mcp_scanner import FakeControl, _test_screen, _semantic_result
    button=FakeControl("Save","ButtonControl",(40,30,160,90))
    window=FakeControl("Example","WindowControl",(0,0,400,300),[button])
    def xpath(control):
        return [dict(ControlType="WindowControl",Name="Example"),
                dict(ControlType=control.ControlTypeName,Name=control.Name)]
    monkeypatch.setattr(scanner,"_activate_window",lambda _:None)
    monkeypatch.setattr(scanner,"_find_window",lambda _:window)
    monkeypatch.setattr(scanner,"_process_name",lambda _:"example-app")
    monkeypatch.setattr(scanner,"_screenshot_window",lambda _:_test_screen())
    monkeypatch.setattr(scanner,"segment_interface",lambda *a:dict(page=dict(id="main",name="Main"),regions=[]))
    monkeypatch.setattr(scanner,"get_control_xpath",xpath)
    monkeypatch.setattr(scanner,"_verify_location",lambda _location,rect,*a:(True,rect,"matched"))
    monkeypatch.setattr(scanner,"analyze_control_semantics",lambda image,candidates,*a:_semantic_result(candidates))
    library=KnowledgeLibrary(tmp_path)
    imported=library.upsert_control_record(dict(LOCATION=location_from_xpath(xpath(button)),app_name="Example"),
                                          app_id="example-app",page_id="main")
    result=scanner.scan_window("Example","unused","unused","unused","1.0.2",strategy="full-uia",root=tmp_path,show_overlay=False)
    buttons=[r for r in library.read("example-app")["controls"] if r.get("control_type")=="ButtonControl"]
    assert len(buttons)==1 and buttons[0]["id"]==imported["id"]
    assert imported["id"] in result["observed_control_ids"]
