"""Agent/GUI integration using real vault files and fake workers, never desktop input."""
from copy import deepcopy
import json
from pathlib import Path
import threading

import pytest

from easy_uiauto import workflows
from easy_uiauto.knowledge_library import PAGE_ID
from easy_uiauto.mcp import knowledge, library_tools, server
from test_knowledge_library import add_control, library
from test_region_picker import picked


@pytest.fixture
def window(monkeypatch):
    value = dict(title="Test", class_name="Window", handle=123, process_id=456, visible=True)
    monkeypatch.setattr(library_tools.native_windows, "matching_windows", lambda *_a, **_kw: [value])
    return value


@pytest.fixture
def region(library, tmp_path):
    identifier = library.save_picked_region("demo", picked(tmp_path))
    library.configure_region("demo", identifier, steps=[], description="用户指定：工程操作入口")
    return identifier


def definition(**overrides):
    return dict(keyword="打开工程", steps=[dict(id="s1", action="click",
        target=dict(label="打开"), checks=[])], **overrides)


def test_read_manual_records_and_bounded_documents_in_selected_vault(library, region):
    control = add_control(library)
    # Save the current assignment as it would be produced by learning.
    control.update(region_id=region.removeprefix(PAGE_ID+"."), description="用户手动写的说明")
    knowledge.save_control(library.directory("demo"), control)
    reference = dict(id="ref-one", kind="reference", app_id="demo", name="参考", enabled=False)
    knowledge.write_markdown(library.directory("demo")/"references/ref-one.md", reference, "abcdefghij")
    result = library_tools.get_library("demo", vault=str(library.root), section="controls", region_id=region)
    assert result["vault"] == str(library.root)
    assert result["records"][0]["description"] == control["description"]
    assert result["records"][0]["effective_entry_steps"] == []
    assert library_tools.get_library("demo", vault=str(library.root), section="regions")["records"][0]["description"].startswith("用户指定")
    listing = library_tools.get_library("demo", vault=str(library.root), section="references")["records"][0]
    assert listing["enabled"] is False and "body" not in listing
    body = library_tools.get_library("demo", vault=str(library.root), section="references",
                                    record_id="ref-one", text_offset=2, text_limit=3)["records"][0]
    assert body["body"] == "cde" and body["next_text_offset"] == 5
    listing = library_tools.get_library(vault=str(library.root), limit=1)
    assert listing["records"][0]["id"] == "demo" and listing["next_offset"] is None
    overview = library_tools.get_library("demo", vault=str(library.root))
    assert overview["counts"]["controls"] == 1


def test_record_pagination_and_unknown_scope_errors(library):
    control = add_control(library)
    knowledge.save_control(library.directory("demo"), {**control, "id": "control-two"})
    first = library_tools.get_library("demo", vault=str(library.root), section="controls", limit=1)
    second = library_tools.get_library("demo", vault=str(library.root), section="controls", offset=first["next_offset"], limit=1)
    assert first["records"][0]["id"] != second["records"][0]["id"]
    assert first["total"] == 2 and second["next_offset"] is None
    for kwargs in [dict(limit=0), dict(section="unknown"), dict(section="controls", record_id="missing"),
                   dict(section="controls", region_id="missing")]:
        with pytest.raises(ValueError): library_tools.get_library("demo", vault=str(library.root), **kwargs)


def test_prepare_learning_retains_selected_manual_region_and_state(library, region, window):
    before = library.read("demo")
    request = library_tools._request("learn-region", "demo", vault=str(library.root), region_id=region)
    assert request["operation"] == "learn" and request["region_id"] == region
    assert request["layout"]["app_id"] == "demo" and request["layout"]["page_id"] == PAGE_ID
    assert request["layout"]["restrict_regions"] is True
    assert request["layout"]["regions"][0]["id"] == region.removeprefix(PAGE_ID+".")
    assert request["layout"]["regions"][0]["normalized_rect"] == before["regions"][0]["normalized_rect"]
    assert request["semantic_mode"] == "native" and request["capture_handshake"] is False
    assert library.read("demo") == before  # Preparing never rewrites the manual layout.


def test_other_page_learning_uses_that_pages_namespace(library, tmp_path, window):
    scoped = library.scoped("another-state")
    identifier = scoped.save_picked_region("demo", picked(tmp_path))
    request = library_tools._request("learn-region", "demo", vault=str(library.root), region_id=identifier)
    assert request["layout"]["page_id"] == "another-state"


def test_reject_unconfirmed_and_dynamic_regions_and_ambiguous_window(library, region, window, monkeypatch):
    page = library.page("demo")
    knowledge.save_page(library.directory("demo"), {**page, "layout_confirmed": False})
    with pytest.raises(ValueError, match="确认"):
        library_tools._request("learn-region", "demo", vault=str(library.root), region_id=region)
    saved = library.regions("demo")[0]
    knowledge.save_region(library.directory("demo"), {**saved, "source": "dynamic-container"})
    with pytest.raises(ValueError, match="expand-control"):
        library_tools._request("learn-region", "demo", vault=str(library.root), region_id=region)
    monkeypatch.setattr(library_tools.native_windows, "matching_windows", lambda *_a, **_kw: [window, window])
    with pytest.raises(ValueError, match="唯一"):
        library_tools._request("learn-region", "demo", vault=str(library.root), region_id=region)


def test_expand_and_reference_supplement_share_gui_payloads(library, region, window, monkeypatch):
    control = add_control(library)
    control["region_id"] = region.removeprefix(PAGE_ID+".")
    knowledge.save_control(library.directory("demo"), control)
    expanded = library_tools._request("expand-control", "demo", vault=str(library.root), control_id=control["id"])
    assert expanded["operation"] == "expand" and expanded["layout"]["page_id"] == PAGE_ID
    assert expanded["control_id"] == control["id"]
    monkeypatch.setattr(library_tools.native_windows, "matching_windows", lambda *_a, **_kw: pytest.fail("No desktop for supplementation"))
    supplement = library_tools._request("supplement-descriptions", "demo", vault=str(library.root), region_id=region)
    assert supplement["control_ids"] == [control["id"]] and "window" not in supplement


def test_save_and_idempotently_update_real_gui_workflow(library):
    add_control(library)
    before = deepcopy(library.read("demo")["controls"])
    result = library_tools.save_workflow("demo", definition(), vault=str(library.root))
    assert result["action"] == "created" and result["verified"] is False
    assert Path(result["document_path"]).is_file()
    saved = library.workflows("demo")[0]
    assert saved["status"] == "ready" and saved["steps"][0]["target"]["control_id"] == "control-one"
    updated = library_tools.save_workflow("demo", definition(name="自定义名称"), vault=str(library.root))
    assert updated["action"] == "updated" and updated["workflow"]["id"] == saved["id"]
    assert len(library.workflows("demo")) == 1 and library.read("demo")["controls"] == before


def test_ambiguous_and_missing_targets_stay_draft(library):
    control = add_control(library)
    knowledge.save_control(library.directory("demo"), {**control, "id": "duplicate"})
    result = library_tools.save_workflow("demo", definition(), vault=str(library.root))
    assert result["workflow"]["status"] == "draft"
    assert not result["workflow"]["steps"][0]["target"].get("control_id")
    flow = definition()
    flow["steps"][0]["target"] = {"control_id": ""}  # No label, never crash or guess.
    assert library_tools.save_workflow("demo", flow, vault=str(library.root))["workflow"]["status"] == "draft"


def test_empty_target_label_never_binds_to_an_unnamed_control(library):
    control = add_control(library)
    control.update(name="", semantic_name="", aliases=[])
    knowledge.save_control(library.directory("demo"), control)
    flow = definition()
    flow["steps"][0]["target"] = {"label": ""}
    saved = library_tools.save_workflow("demo", flow, vault=str(library.root))["workflow"]
    assert saved["status"] == "draft" and not saved["steps"][0]["target"].get("control_id")


@pytest.mark.parametrize("change", [dict(last_run={"ok": True}), dict(status="verified"), dict(id="invented")])
def test_reject_fabricated_evidence_or_unknown_workflow(library, change):
    with pytest.raises(ValueError):
        library_tools.save_workflow("demo", {**definition(), **change}, vault=str(library.root))
    assert library.workflows("demo") == []


def test_reject_foreign_control_and_preserve_old_execution_as_outdated(library):
    control = add_control(library)
    flow = definition()
    flow["steps"][0]["target"]["control_id"] = "foreign-control"
    with pytest.raises(ValueError, match="不属于"):
        library_tools.save_workflow("demo", flow, vault=str(library.root))
    saved = library_tools.save_workflow("demo", definition(), vault=str(library.root))["workflow"]
    library.save_workflow_execution("demo", saved["id"], dict(ok=True,
        definition_hash=workflows.definition_hash(saved), control_hash=workflows.control_hash(saved, [control])))
    assert not library.workflows("demo")[0]["last_run"]["outdated"]
    revised = definition()
    revised["steps"][0]["action"] = "double-click"
    result = library_tools.save_workflow("demo", revised, vault=str(library.root))["workflow"]
    assert result["last_run"]["ok"] and result["last_run"]["outdated"]


def test_run_requires_ready_flow_and_resolves_parameters_before_worker(library, window):
    saved = library_tools.save_workflow("demo", definition(), vault=str(library.root))["workflow"]
    with pytest.raises(ValueError, match="完善"):
        library_tools._request("run-workflow", "demo", vault=str(library.root), workflow_id=saved["id"])
    add_control(library)
    flow = definition()
    flow["steps"][0].update(action="set-text", text="${工程名}")
    saved = library_tools.save_workflow("demo", flow, vault=str(library.root))["workflow"]
    with pytest.raises(ValueError, match="参数"):
        library_tools._request("run-workflow", "demo", vault=str(library.root), workflow_id=saved["id"])
    request = library_tools._request("run-workflow", "demo", vault=str(library.root), workflow_id=saved["id"], parameters={"工程名": "演示"})
    assert request["operation"] == "workflow" and request["flow_id"] == saved["id"]
    assert request["parameters"] == {"工程名": "演示"}


class Worker:
    def __init__(self, payload, *, timeout):
        self.payload, self.timeout = payload, timeout
        self.folder = Path("fake-worker")
        self.observed = threading.Event()
        self.state = dict(state="running", worker_running=True, message="learning")
        self.cancelled = self.closed = 0

    def poll(self):
        self.observed.set()
        return deepcopy(self.state)

    def cancel(self):
        self.cancelled += 1
        self.state = dict(state="cancelled", worker_running=False,
                          execution=dict(input_attempted=True, steps=[dict(completed=True)]))

    def close(self):
        self.closed += 1


@pytest.fixture
def manager(monkeypatch):
    monkeypatch.setattr(library_tools, "LibraryTask", Worker)
    monkeypatch.setattr(library_tools, "_request", lambda *_a, **_kw: dict(vault="test-vault"))
    value = library_tools.LibraryJobs()
    monkeypatch.setattr(library_tools, "jobs", value)
    yield value
    value.close()
    if value.thread: value.thread.join(2)


def test_worker_outlives_start_and_keeps_result_after_cleanup(manager):
    started = manager.start("learn-region", "demo")
    worker = manager.jobs[started["task_id"]]["worker"]
    assert started["state"] == "running" and worker.observed.wait(2)
    with pytest.raises(ValueError, match="不要重复"):
        manager.start("learn-region", "demo")
    with manager.lock:
        worker.state = dict(state="completed", worker_running=True, result=dict(ok=False, message="failed check"))
    # A published result is not terminal until the process has exited.
    with pytest.raises(ValueError): manager.start("learn-region", "demo")
    with manager.lock: worker.state["worker_running"] = False
    manager.thread.join(2)
    result = manager.status()
    assert result["state"] == "completed" and not result["ok"]
    assert result["result"]["message"] == "failed check"
    assert worker.closed == 1 and manager.active == ""
    assert "worker" not in result and "worker" not in manager.jobs[started["task_id"]]


def test_cancel_preserves_partial_input_evidence_and_releases_worker_once(manager):
    started = manager.start("expand-control", "demo")
    worker = manager.jobs[started["task_id"]]["worker"]
    assert manager.status(stop=True)["state"] == "cancelling"
    manager.thread.join(2)
    result = manager.status()
    assert result["state"] == "cancelled" and not result["ok"]
    assert result["execution"]["input_attempted"] and worker.closed == 1
    manager.close()
    assert worker.closed == 1


def test_monitor_handles_worker_failure_without_another_status_request(manager):
    started = manager.start("supplement-descriptions", "demo")
    worker = manager.jobs[started["task_id"]]["worker"]
    assert worker.timeout == 3600
    with manager.lock:
        worker.state = dict(state="failed", worker_running=False, error="timeout")
    manager.thread.join(2)
    assert manager.status()["error"] == "timeout" and worker.closed == 1


def test_poll_exception_closes_worker_and_reports_failure(manager):
    started = manager.start("learn-region", "demo")
    worker = manager.jobs[started["task_id"]]["worker"]
    def fail(): raise OSError("unreadable state")
    with manager.lock: worker.poll = fail
    manager.thread.join(2)
    assert manager.status()["state"] == "failed" and worker.closed == 1
    assert manager.status()["error"] == "unreadable state"


def test_close_during_poll_does_not_leave_active_job_or_double_close(manager):
    started = manager.start("learn-region", "demo")
    worker = manager.jobs[started["task_id"]]["worker"]
    manager.close()
    manager.thread.join(2)
    assert manager.status()["state"] == "cancelled" and not manager.active
    assert worker.closed == 1


def test_mcp_mode_guard_and_json_errors(manager, monkeypatch):
    monkeypatch.setattr(server, "CURRENT_MODE", "learn")
    result = json.loads(server.start_ui_library_task("learn-region", "demo"))
    assert not result["ok"] and "blocked" in result["error"] and not manager.jobs
    assert json.loads(server.start_ui_library_task("supplement-descriptions", "demo"))["ok"]
    assert not json.loads(server.get_ui_library_task("missing"))["ok"]


def test_readiness_is_local_and_never_returns_api_key(monkeypatch):
    from easy_uiauto import image_models
    from easy_uiauto.mcp import configuration
    monkeypatch.setattr(image_models, "inspect_models", lambda: dict(ok=True, models=[], runtime_issues=[]))
    monkeypatch.setattr(configuration, "load_gui_configuration", lambda: dict(api_url="url", api_key="secret", api_model="model"))
    result = server.get_ui_library_readiness()
    assert json.loads(result)["ready"] and json.loads(result)["reference_api_configured"]
    assert "secret" not in result
