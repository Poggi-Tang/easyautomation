"""Cross-region expectations use bounded container queries and runtime values."""
from copy import deepcopy
from types import SimpleNamespace

import pytest

from easy_uiauto import workflows
from easy_uiauto.mcp.results import ActionError
from test_workflows import runtime


def node(name,*children):return SimpleNamespace(Name=name,GetChildren=lambda:list(children))


def test_dynamic_node_checks_are_scoped_and_missing_is_not_guessed_after_limit():
    tree=node("工程树",node("Root",node("Demo_001")))
    assert workflows._check(dict(kind="node-exists",expected="Demo_001"),tree)["ok"]
    assert workflows._check(dict(kind="node-absent",expected="Other"),tree)["ok"]
    assert not workflows._check(dict(kind="node-absent",expected="Demo_001"),tree)["ok"]
    cyclic=node("Cycle");cyclic.GetChildren=lambda:[cyclic]
    with pytest.raises(ValueError,match="上限"):
        workflows._check(dict(kind="node-absent",expected="missing"),cyclic)


def test_missing_expectation_parameter_stops_before_any_input(runtime):
    flow,controls,events,_=runtime
    flow["steps"][0]["checks"][0]["expected"]=["${required_name}"]
    with pytest.raises(ValueError,match="缺少流程参数"):workflows.execute(flow,controls,dict(handle=1))
    assert not events


def test_not_visible_wait_accepts_closed_dialog_but_not_invalid_locator(runtime,monkeypatch):
    from easy_uiauto import library_actions
    flow,controls,events,objects=runtime
    check=flow["steps"][0]["checks"][0]
    check.update(kind="not-visible")
    target_id=check["target"]["control_id"]
    original=library_actions.resolve_record
    def resolve(record,*args,**kwargs):
        if record["id"]==target_id:raise ActionError("control_not_found","dialog closed")
        return original(record,*args,**kwargs)
    monkeypatch.setattr(library_actions,"resolve_record",resolve)
    result=workflows.execute(flow,controls,dict(handle=1))
    assert result["steps"][0]["checks"][0]["visible"] is False
    assert events.count("right-click")==events.count("click")==1


def test_runtime_parameters_expand_expectations_without_changing_definition():
    original=dict(steps=[dict(text="${name}",checks=[dict(expected="${name}")])])
    before=deepcopy(original)
    expanded=workflows.substitute(original,{"name":"工程A"})
    assert expanded["steps"][0]["checks"][0]["expected"]=="工程A" and original==before
