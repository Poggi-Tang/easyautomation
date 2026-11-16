"""Locator checks survive persistence without approving business semantics."""
from copy import deepcopy
import json

import pytest

from easy_uiauto import library_tasks, locator_verification, workflows
from easy_uiauto.mcp import knowledge
from easy_uiauto.mcp.protocol import location_from_xpath
from test_knowledge_library import library, add_control
from test_gui_image_replay import visual


def image_record(library, *, quarantine=None):
    record = add_control(library, status="quarantined" if quarantine else "observed", quarantine=quarantine)
    record.update(control_type="VisualElement", location=location_from_xpath([]), image="images/component.png",
                  semantic_status="uncertain", function_verification={"status":"inferred", "executed":False})
    knowledge.save_control(library.directory("demo"),record)
    return record


def checked(record, **values):
    return dict(ok=True, resolved_by="image", locator_signature=locator_verification.signature(record), **values)


@pytest.mark.parametrize("method", ["image", "location"])
def test_check_persists_separately_from_semantics_and_survives_rescan(library, method):
    record=image_record(library)
    result={**checked(record),"resolved_by":method}
    assert library.save_locator_verification("demo",record["id"],result)
    saved=library.read("demo")["controls"][0]
    assert locator_verification.current(saved)["resolved_by"]==method
    assert saved["status"]==record["status"] and saved["semantic_status"]=="uncertain"
    assert saved["function_verification"]==record["function_verification"]
    # A scan of the same locator must not erase the independent runtime check.
    knowledge.save_control(library.directory("demo"),{**record,"scan_id":"scan-again"})
    saved=knowledge.find_control_record(library.directory("demo"),record["id"])[1]
    assert locator_verification.current(saved)["status"]=="passed"


@pytest.mark.parametrize("change", ["location", "image", "window_scope"])
def test_edits_invalidate_previous_evidence_and_reject_late_results(library,change):
    record=image_record(library)
    result=checked(record)
    assert library.save_locator_verification("demo",record["id"],result)
    edited=deepcopy(library.read("demo")["controls"][0])
    edited[change]={"kind":"owned","class_name":"Popup"} if change=="window_scope" else (
        location_from_xpath([{"ControlType":"ButtonControl","Name":"Other"}]) if change=="location" else "images/other.png")
    knowledge.save_control(library.directory("demo"),edited)
    assert not locator_verification.current(edited)
    assert not library.save_locator_verification("demo",record["id"],result)


def test_later_failed_check_replaces_success_without_releasing_quarantine(library):
    record=image_record(library,quarantine={"active":True,"reason":"runtime-failure"})
    assert library.save_locator_verification("demo",record["id"],checked(record))
    failure={**checked(record),"ok":False,"error":"image no longer matches"}
    assert library.save_locator_verification("demo",record["id"],failure)
    saved=library.read("demo")["controls"][0]
    assert locator_verification.current(saved)["status"]=="failed" and knowledge.runtime_quarantined(saved)


@pytest.mark.parametrize("save_as_flow", [False, True])
def test_image_execution_updates_the_same_control_record(library,visual,save_as_flow):
    record=image_record(library)
    record["image"]=str(visual.directory/"button.png")
    knowledge.save_control(library.directory("demo"),record)
    flow=dict(id="flow",keyword="元件操作",steps=[workflows.click_step(record)])
    if save_as_flow:library.save_workflow("demo",flow)
    result=workflows.execute(flow,[record],visual.expected,directory=library.directory("demo"))
    assert result["ok"] and result["steps"][0]["resolved_by"]=="image"
    if save_as_flow:library.save_workflow_execution("demo","flow",result)
    else:library.save_execution("demo",record["id"],result)
    saved=library.read("demo")["controls"][0]
    assert locator_verification.current(saved)["status"]=="passed"
    assert saved["function_verification"]["executed"] is False and saved["status"]=="observed"


@pytest.mark.parametrize("step", [dict(skipped=True,ok=True),dict(completed=False,ok=False),dict(control_id="another",ok=True)])
def test_skipped_failed_or_unrelated_steps_do_not_verify_selected_target(library,step):
    record=image_record(library)
    evidence={**checked(record),"control_id":record["id"],"ok":False,**step}
    library.save_execution("demo",record["id"],dict(ok=True,steps=[evidence]))
    assert not locator_verification.current(library.read("demo")["controls"][0])


def test_verify_worker_records_success_then_failure_without_input(library,visual,tmp_path):
    record=image_record(library)
    record["image"]=str(visual.directory/"button.png")
    knowledge.save_control(library.directory("demo"),record)
    folder=tmp_path/"verify-job";folder.mkdir()
    library_tasks._write(folder/"request.json",dict(operation="verify",vault=str(library.root),
        control_id=record["id"],window=visual.expected,**record))
    library_tasks._worker(folder)
    state=json.loads((folder/"state.json").read_text(encoding="utf-8"))
    assert state["state"]=="completed" and state["result"]["verification_saved"]
    assert locator_verification.current(library.read("demo")["controls"][0])["status"]=="passed"
    visual.screen.paste("white",(0,0,240,160))
    library_tasks._worker(folder)
    state=json.loads((folder/"state.json").read_text(encoding="utf-8"))
    assert state["state"]=="failed"
    assert locator_verification.current(library.read("demo")["controls"][0])["status"]=="failed"
    assert not any(isinstance(event,tuple) for event in visual.events)
