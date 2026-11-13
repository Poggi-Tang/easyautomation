"""Shared workflow references must resolve before input and invalidate old evidence."""
import pytest

from easy_uiauto import workflows
from test_knowledge_library import library, add_control
from test_workflows import runtime


def definitions(record):
    child=dict(id="shared",keyword="公共填写",parameters={"name":"default"},steps=[
        {**workflows.click_step(record,action="set-text"),"text":"${name}"}])
    parent=dict(id="entry",keyword="菜单新建",parameters={"name":"parent"},steps=[
        dict(id="call",action="call",workflow_id="shared",parameters={"name":"${project}"})])
    return child,parent


def test_call_parameters_and_ids_are_resolved_without_mutating_definitions(library):
    record=add_control(library);child,parent=definitions(record)
    expanded=workflows.expand_calls(parent,{"shared":child},parameters={"project":"demo"})
    assert expanded["steps"][0]["text"]=="demo" and expanded["steps"][0]["id"].startswith("call/")
    assert child["steps"][0]["text"]=="${name}" and parent["steps"][0]["action"]=="call"


def test_missing_and_cyclic_calls_are_not_ready(library):
    record=add_control(library);child,parent=definitions(record)
    assert not workflows.readiness(parent,[record])["ready"]
    child["steps"]=[dict(id="back",action="call",workflow_id="entry")]
    report=workflows.readiness(parent,[record],{"shared":child,"entry":parent})
    assert not report["ready"] and "循环" in report["issues"][0]["message"]


def test_shared_edit_invalidates_parent_run_and_deleted_child_is_draft(library):
    record=add_control(library);child,parent=definitions(record)
    library.save_workflow("demo",child)
    parent=library.save_workflow("demo",parent)
    catalog=workflows.load_catalog(library.directory("demo"))
    parent["last_run"]=dict(definition_hash=workflows.definition_hash(parent,catalog),
                            control_hash=workflows.control_hash(parent,[record],catalog))
    library.save_workflow("demo",parent)
    child["steps"][0]["text"]="changed";library.save_workflow("demo",child)
    saved=next(r for r in library.workflows("demo") if r["id"]=="entry")
    assert saved["last_run"]["outdated"]
    (library.directory("demo")/"workflows/shared.md").unlink()
    saved=next(r for r in library.workflows("demo") if r["id"]=="entry")
    assert saved["status"]=="draft" and saved["last_run"]["outdated"]


def test_execute_nested_workflow_keeps_assertions_and_rejects_missing_child_before_input(runtime):
    flow,controls,events,_=runtime
    parent=dict(id="parent",keyword="入口",steps=[dict(id="shared",action="call",workflow_id="child")])
    with pytest.raises(ValueError,match="不存在"):
        workflows.execute(parent,controls,dict(handle=1),catalog={})
    assert not events
    result=workflows.execute(parent,controls,dict(handle=1),catalog={"child":flow})
    assert result["ok"] and all(s["id"].startswith("shared/") for s in result["steps"])
    assert events.count("click")==events.count("right-click")==1
    assert sum(len(s["checks"]) for s in result["steps"])==7


def test_expansion_rejects_colliding_flat_ids(library):
    record=add_control(library);child,_=definitions(record)
    child["steps"][0]["id"]="step"
    parent=dict(keyword="collision",steps=[dict(id="call",action="call",workflow_id="shared"),
                {**workflows.click_step(record),"id":"call/step"}])
    with pytest.raises(ValueError,match="ID 重复"):workflows.expand_calls(parent,{"shared":child})
