"""Limited UIA exposure must not silently masquerade as a missing control."""
from copy import deepcopy
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from easy_uiauto import library_actions as actions
from easy_uiauto.mcp.protocol import location_from_xpath
from easy_uiauto.mcp.results import ActionError


def denied():
    error=OSError("Access is denied");error.winerror=5
    return error


@pytest.mark.parametrize("wrapped",[False,True])
def test_direct_and_wrapped_com_denial_report_permission_without_guessing_elevation(wrapped):
    cause=RuntimeError("COM access denied");cause.hresult=-2147024891
    error=ActionError("action_failed","UIA lookup failed") if wrapped else cause
    if wrapped:error.__cause__=cause
    result=actions.diagnose_uia_error(error,expected=dict(title="SimuLab"))
    assert result.code=="operation_blocked" and result.reason=="uia_access_denied"
    assert "Windows 拒绝" in str(result) and "相同的账户和权限" in str(result)
    assert "普通权限" not in str(result)


@pytest.fixture
def limited_root(monkeypatch):
    root=SimpleNamespace(ClassName="Qt5152QWindowIcon")
    record=dict(name="CommonToolBar",location=location_from_xpath([
        dict(ControlType="WindowControl",Name="SimuLab",ClassName="SimuUI"),
        dict(ControlType="ToolBarControl",Name="CommonToolBar")]))
    expected=dict(process_id=123,title="SimuLab")
    activation=Mock(side_effect=AssertionError("Diagnostics activated the window"))
    monkeypatch.setattr(actions.Controller,"activate_window_handle",activation)
    return root,record,expected


def test_query_denial_is_reported_as_unknown_not_proof_of_admin(limited_root,monkeypatch):
    root,record,expected=limited_root
    query=Mock(side_effect=[False,denied()]);monkeypatch.setattr(actions.native_windows,"process_elevated",query)
    error=ActionError("control_not_found","Control not found",reason="index_out_of_range")
    result=actions.diagnose_uia_error(error,record=record,expected=expected,root=root)
    assert result.code==error.code and result.reason==error.reason
    assert "SimuUI" in result.message and "Qt5152QWindowIcon" in result.message
    assert "无权读取目标进程" in result.message and "若目标软件" in result.message
    assert "目标软件以管理员身份运行，管理器为普通权限" not in result.message


def test_confirmed_integrity_difference_blocks_lookup_with_existing_record_untouched(limited_root,monkeypatch):
    root,record,expected=limited_root;before=deepcopy(record)
    monkeypatch.setattr(actions.native_windows,"process_elevated",Mock(side_effect=[False,True]))
    monkeypatch.setattr(actions,"observe_control",Mock(side_effect=ActionError("control_not_found","Control not found")))
    with pytest.raises(ActionError) as failure:actions.resolve_record(record,expected,root)
    assert failure.value.code=="operation_blocked" and failure.value.reason=="uia_access_denied"
    assert "管理器为普通权限" in str(failure.value) and record==before


def test_real_missing_control_under_matching_root_keeps_error_and_skips_privilege_probe(limited_root,monkeypatch):
    root,record,expected=limited_root;root.ClassName="SimuUI"
    query=Mock(side_effect=AssertionError("Ordinary lookup should not query other processes"))
    monkeypatch.setattr(actions.native_windows,"process_elevated",query)
    error=ActionError("control_not_found","Control not found",retryable=True)
    assert actions.diagnose_uia_error(error,record=record,expected=expected,root=root) is error
    query.assert_not_called()


@pytest.mark.parametrize("code,reason",[("invalid_argument",""),("ambiguous_match",""),("window_invalid","window_identity_mismatch")])
def test_other_locator_failures_are_not_reinterpreted(limited_root,monkeypatch,code,reason):
    root,record,expected=limited_root
    monkeypatch.setattr(actions.native_windows,"process_elevated",Mock(side_effect=AssertionError("Unexpected probe")))
    error=ActionError(code,"Invalid locator",reason=reason)
    assert actions.diagnose_uia_error(error,record=record,expected=expected,root=root) is error


def test_failed_diagnostic_cannot_replace_original_reason(limited_root,monkeypatch):
    root,record,expected=limited_root
    monkeypatch.setattr(actions.native_windows,"process_elevated",Mock(side_effect=RuntimeError("probe unavailable")))
    error=ActionError("control_not_found","Control not found",reason="position_context_missing")
    result=actions.diagnose_uia_error(error,record=record,expected=expected,root=root)
    assert result.code==error.code and result.reason==error.reason
    assert "probe unavailable" not in result.message


def test_file_access_denial_is_not_relabeled_as_uia_permission_failure():
    error=denied()
    assert actions.diagnose_uia_error(error) is error
    assert actions.diagnose_uia_error(error,uia_operation=True).reason=="uia_access_denied"


def test_worker_publishes_structured_denial_and_preserves_partial_execution(tmp_path,monkeypatch):
    import json
    import uiautomation
    from easy_uiauto import library_tasks
    request=dict(operation="verify",location={},window=dict(handle=1,title="SimuLab"))
    (tmp_path/"request.json").write_text(json.dumps(request),encoding="utf-8")
    cause=RuntimeError("COM denied");cause.hresult=-2147024891
    error=ActionError("action_failed","UIA lookup failed");error.__cause__=cause
    error.execution_result=dict(input_attempted=True,steps=[dict(completed=True)])
    monkeypatch.setattr(actions,"bind_window",Mock(side_effect=error))
    library_tasks._run_worker(tmp_path)
    state=json.loads((tmp_path/"state.json").read_text(encoding="utf-8"))
    assert state["state"]=="failed" and state["error_code"]=="operation_blocked"
    assert state["error_reason"]=="uia_access_denied" and "Windows 拒绝" in state["error"]
    assert state["execution"]==error.execution_result
