"""Relearning scenarios use a deterministic app and never send desktop input."""
from copy import deepcopy
import json

import pytest

from easy_uiauto import relearning, workflows
from easy_uiauto.deep_learning import Explorer
from easy_uiauto.mcp import knowledge
from test_deep_learning import Backend
from test_knowledge_library import library, add_control
from test_region_picker import picked


@pytest.fixture
def region(library,tmp_path):
    library.save_picked_region("demo",picked(tmp_path))
    return library.regions("demo")[0]


def test_continue_replans_blocked_descendant_after_inputs_and_references_change(library,region):
    first=Explorer(library,"demo",region,Backend(library,blocked=True),inputs={})
    assert not first.run()["complete"]
    library.save_usage("demo","补充资料","填写工程信息后点击完成")
    backend=Backend(library)
    second=Explorer(library,"demo",region,backend,inputs={"工程路径":"D:/test"},instructions="按补充资料学习")
    result=second.run()
    assert result["complete"] and second.graph["id"]==first.graph["id"]
    assert backend.executed==["menu","fill"]
    assert backend.planned==["root","dialog","done"]
    assert len(list((library.directory("demo")/"learning/history").glob("*.json")))==1


def test_continue_retries_failure_before_input_but_keeps_uncertain_input(library,region):
    backend=Backend(library)
    def fail(steps):
        error=ValueError("定位失败，未发送输入");error.execution_result=dict(input_attempted=False)
        raise error
    backend.execute=fail
    first=Explorer(library,"demo",region,backend);first.run()
    assert first.graph["states"]["root"]["branches"][0]["status"]=="failed"
    second=Explorer(library,"demo",region,Backend(library))
    assert second.run()["complete"]


def test_user_can_explicitly_retry_uncertain_branch_without_forking_session(library,region):
    first=Explorer(library,"demo",region,Backend(library,fail=True))
    with pytest.raises(InterruptedError):first.run()
    branch=first.graph["states"]["root"]["branches"][0]["id"]
    backend=Backend(library)
    second=Explorer(library,"demo",region,backend,retry_branches=[branch])
    assert second.run()["complete"] and backend.executed[0]=="menu"
    assert second.graph["id"]==first.graph["id"]


@pytest.mark.parametrize("field,value",[("text","保留人工参数"),("description","人工调整的步骤")])
def test_improve_retains_workflow_ids_and_human_edits(library,region,field,value):
    first=Explorer(library,"demo",region,Backend(library));first.run()
    before=workflows.load_catalog(library.directory("demo"))
    shared=next(f for f in before.values() if f["steps"][0]["target"]["control_id"]=="fill")
    edited=deepcopy(shared);edited["steps"][0][field]=value
    library.save_workflow("demo",edited)
    backend=Backend(library)
    result=Explorer(library,"demo",region,backend,mode="improve").run()
    after=workflows.load_catalog(library.directory("demo"))
    assert set(after)==set(before)
    assert after[shared["id"]]["steps"][0][field]==value
    assert result["complete"]


def test_improve_selected_control_does_not_execute_other_entries(library,region):
    first=Explorer(library,"demo",region,Backend(library));first.run()
    backend=Backend(library)
    Explorer(library,"demo",region,backend,mode="improve",control_ids=["toolbar"]).run()
    assert backend.executed==["toolbar"]


def test_improve_cannot_claim_success_for_unreached_selection(library,region):
    backend=Backend(library)
    result=Explorer(library,"demo",region,backend,mode="improve",control_ids=["control-one"]).run()
    assert not result["complete"] and not backend.executed
    assert any("尚未到达" in issue["reason"] for issue in result["issues"])


def test_changed_suffix_updates_original_entry_flow_in_place(library,region):
    explorer=Explorer(library,"demo",region,Backend(library));explorer.run()
    original=workflows.load_catalog(library.directory("demo"))
    roots=set(explorer.graph["entry_workflow_ids"])
    shared=next(f for f in original.values() if f["steps"][0].get("target",{}).get("control_id")=="fill")
    result=deepcopy(explorer.graph["states"]["done"]);result["id"]="new-result"
    explorer.graph["states"]["new-result"]=result
    explorer.graph["states"]["dialog"]["branches"][0]["target_state"]="new-result"
    explorer.archive()
    current=workflows.load_catalog(library.directory("demo"))
    assert set(explorer.graph["entry_workflow_ids"])==roots
    assert current[shared["id"]]["steps"][-1]["workflow_id"]!=shared["steps"][-1]["workflow_id"]


def test_continue_at_saved_result_does_not_repeat_parent_submission(library,region):
    backend=Backend(library)
    plan=backend.plan
    def with_commit(observation,route):
        result=plan(observation,route)
        if observation["state_id"]=="root":
            for b in result["branches"]:b.update(replayable=False,effect="commit")
        return result
    backend.plan=with_commit
    first=Explorer(library,"demo",region,backend,max_actions=1);first.run()
    second_backend=Backend(library);second_backend.current="dialog"
    second=Explorer(library,"demo",region,second_backend)
    second.run()
    assert second_backend.executed==["fill"]
    assert second.graph["states"]["root"]["branches"][0]["status"]=="done"


def test_session_selection_is_scoped_and_new_mode_is_explicit(library,region):
    first=Explorer(library,"demo",region,Backend(library));first.run()
    with pytest.raises(ValueError,match="不属于"):
        Explorer(library,"demo",{**region,"id":"another"},Backend(library),session_id=first.graph["id"])
    new=Explorer(library,"demo",region,Backend(library),mode="new")
    assert new.graph["id"]!=first.graph["id"]
    assert relearning.select_session(library.directory("demo"),region["id"])["id"]==first.graph["id"]


def test_update_marks_missing_not_deleted_and_invalidates_transitive_flow(library,region):
    record=add_control(library)
    record.update(region_id=region["id"],semantic_source="manual",semantic_name="人工名称",description="人工说明")
    knowledge.save_control(library.directory("demo"),record)
    child=library.save_workflow("demo",dict(id="child",keyword="子流程",steps=[workflows.click_step(record)]))
    parent=library.save_workflow("demo",dict(id="parent",keyword="入口",steps=[dict(id="call",action="call",workflow_id="child")]))
    revision=relearning.KnowledgeRevision(library,"demo",region["id"],"update")
    revision.reconcile([], [region["id"]])
    summary=revision.summarize()
    assert summary["unobserved"]==[record["id"]] and set(summary["workflows"])=={"child","parent"}
    current=library.read("demo")["controls"]
    assert len(current)==1 and current[0]["availability"]["state"]=="unobserved"
    continued=relearning.KnowledgeRevision(library,"demo",region["id"],"resume")
    assert continued.summarize()["unobserved"]==[record["id"]]
    assert not workflows.readiness(parent,current,workflows.load_catalog(library.directory("demo")))["ready"]
    changed=deepcopy(record);changed.update(semantic_name="自动名称",description="自动描述")
    changed["location"]["Name"]="new-locator"
    restored=revision.reconcile([changed],[region["id"]])[0]
    assert restored["semantic_name"]=="人工名称" and restored["description"]=="人工说明"
    assert restored["location"]["Name"]=="new-locator" and restored["availability"]["state"]=="observed"
    assert record["id"] in revision.summarize()["updated"]
    assert workflows.readiness(parent,library.read("demo")["controls"],workflows.load_catalog(library.directory("demo")))["ready"]
    history=json.loads(revision.path.read_text(encoding="utf-8"))
    assert history["before_controls"][record["id"]]["description"]=="人工说明"


def test_partial_scan_and_unopened_entries_do_not_mark_missing(library,region):
    record=add_control(library);record["region_id"]=region["id"]
    knowledge.save_control(library.directory("demo"),record)
    revision=relearning.KnowledgeRevision(library,"demo",region["id"],"update")
    revision.reconcile([], [region["id"]],complete=False)
    assert not revision.summarize()["unobserved"]
    record["navigation_path"]=["entry"]
    knowledge.save_control(library.directory("demo"),record)
    revision=relearning.KnowledgeRevision(library,"demo",region["id"],"update")
    revision.reconcile([], [region["id"]])
    assert not revision.summarize()["unobserved"]


def test_update_workflow_requires_real_current_result_checks(library,region):
    record=add_control(library)
    flow=library.save_workflow("demo",dict(id="check",keyword="校验",steps=[workflows.click_step(record)],
        update_validation=dict(status="pending",revision="r")))
    evidence=dict(ok=True,definition_hash=workflows.definition_hash(flow),control_hash=workflows.control_hash(flow,[record]),
                  steps=[dict(checks=[dict(ok=True)])])
    saved=library.save_workflow_execution("demo",flow["id"],{**evidence,"control_hash":"old"})
    assert saved["update_validation"]["status"]=="failed"
    saved=library.save_workflow_execution("demo",flow["id"],evidence)
    assert saved["update_validation"]["status"]=="passed"
    saved["steps"][0]["description"]="changed after verification"
    assert library.save_workflow("demo",saved)["update_validation"]["status"]=="pending"


def test_pending_update_validation_is_carried_to_later_run(library,region):
    record=add_control(library);record["region_id"]=region["id"]
    knowledge.save_control(library.directory("demo"),record)
    library.save_workflow("demo",dict(id="old",keyword="旧流程",steps=[workflows.click_step(record)],
        update_validation=dict(status="pending",revision="previous",reason="待准备测试数据")))
    revision=relearning.KnowledgeRevision(library,"demo",region["id"],"resume")
    revision.reconcile([record],[region["id"]])
    assert revision.summarize()["workflows"]["old"]["status"]=="pending"


def test_update_revalidates_business_flow_even_if_locator_unchanged(library,region):
    record=add_control(library);record["region_id"]=region["id"]
    knowledge.save_control(library.directory("demo"),record)
    library.save_workflow("demo",dict(id="old",keyword="旧流程",steps=[workflows.click_step(record)]))
    revision=relearning.KnowledgeRevision(library,"demo",region["id"],"update")
    revision.reconcile([record],[region["id"]])
    result=revision.summarize()
    assert not result["updated"] and result["workflows"]["old"]["status"]=="pending"
