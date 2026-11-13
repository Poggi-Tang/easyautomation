"""Durable learning sessions and conservative reconciliation of observed knowledge."""
from __future__ import annotations

from copy import deepcopy
import json
from uuid import uuid4

from . import workflows
from .mcp import image_storage, knowledge


MODES = {"resume": "继续未完成学习", "improve": "改进学习结果", "update": "适配软件更新", "new": "首次学习"}


def sessions(directory, region_id):
    result = []
    for path in (directory / "learning").glob("deep-*.json"):
        if image_storage._safe_path(directory, path) is None:continue
        try:
            graph = json.loads(path.read_text(encoding="utf-8"))
            if graph.get("id")==path.stem and graph.get("region_id") == region_id and isinstance(graph.get("states"), dict):
                result.append((path.stat().st_mtime_ns,graph))
        except (ValueError, OSError):continue
    return [graph for _,graph in sorted(result,key=lambda row:(row[1].get("updated_at",""),row[0]),reverse=True)]


def select_session(directory, region_id, session_id=""):
    available = sessions(directory, region_id)
    if session_id:
        selected = next((r for r in available if r["id"] == session_id), None)
        if selected is None:raise ValueError("学习记录不存在或不属于当前区域")
        return deepcopy(selected)
    return deepcopy(available[0]) if available else None


def steps_key(steps, *, anchor=False):
    """Exclude generated IDs/prose, but include actual input and checks in definitions."""
    result = []
    for step in steps:
        target = step.get("target", {})
        item = dict(action=step.get("action"), target=target.get("control_id"), hit=target.get("hit", "center"))
        if not anchor:
            item.update(text=step.get("text"), workflow_id=step.get("workflow_id"), parameters=step.get("parameters"),
                        checks=[{k: c.get(k) for k in ("kind", "target", "expected")} for c in step.get("checks", [])])
        result.append(item)
    return knowledge.stable_id("steps", result)


def branch_match(branch, previous):
    exact = [b for b in previous if steps_key(b.get("steps", [])) == steps_key(branch.get("steps", [])) and branch.get("steps")]
    if len(exact) == 1:return exact[0]
    target = branch.get("control_id") or next((s.get("target", {}).get("control_id") for s in branch.get("steps", [])), None)
    matches = [b for b in previous if target and target == (b.get("control_id") or next(
        (s.get("target", {}).get("control_id") for s in b.get("steps", [])), None))]
    if len(matches) == 1:return matches[0]
    named = [b for b in previous if b["name"] == branch["name"]]
    return named[0] if len(named) == 1 else None


def definition(record):
    return {key: record.get(key) for key in ("location", "window_scope", "region_id", "semantic_name", "description",
            "actions", "control_type", "automation_id", "name", "image_sha256", "normalized_rect")}


def authored_key(flow):
    return knowledge.stable_id("definition", {k: flow.get(k) for k in ("keyword","name","precondition","parameters","steps")})


class KnowledgeRevision:
    def __init__(self, library, app_id, region_id, mode, *, control_ids=(), workflow_ids=()):
        self.library, self.app_id, self.region_id, self.mode = library, app_id, region_id, mode
        self.directory = library.directory(app_id)
        snapshot = library.read(app_id)
        self.before = {r["id"]: deepcopy(r) for r in snapshot["controls"]}
        self.catalog = workflows.load_catalog(self.directory)
        region_ids={region_id,*[r["id"] for r in snapshot["regions"] if r.get("learning_origin_region")==region_id]}
        self.scope_ids={r["id"] for r in snapshot["controls"] if r.get("region_id") in region_ids}
        self.control_ids, self.workflow_ids = set(control_ids), set(workflow_ids)
        if self.control_ids - self.before.keys():raise ValueError("所选控件已不存在")
        if self.workflow_ids - self.catalog.keys():raise ValueError("所选组合动作已不存在")
        for identifier in self.workflow_ids:
            self.control_ids.update(t["control_id"] for _, t in workflows.targets(
                workflows.expand_calls(self.catalog[identifier], self.catalog)))
        self.id = "revision-" + uuid4().hex[:12]
        self.path = self.directory / "learning" / "revisions" / (self.id + ".json")
        self.seen, self.conflicts = set(), []
        self.missing={i for i in self.scope_ids if self.before[i].get("availability",{}).get("state")=="unobserved"
                      and (not self.control_ids or i in self.control_ids)}
        self.affected = {}
        self.write()

    def write(self):
        from .library_tasks import _write
        if image_storage._safe_path(self.directory, self.path) is None:raise ValueError("学习修订路径无效")
        with image_storage._vault_lock(self.library.root):
            self.path.parent.mkdir(parents=True, exist_ok=True)
            _write(self.path, dict(id=self.id, region_id=self.region_id, mode=self.mode, before_controls=self.before,
                                  before_workflows=self.catalog, seen=sorted(self.seen), missing=sorted(self.missing),
                                  conflicts=self.conflicts, affected_workflows=self.affected, updated_at=knowledge.utc_now()))

    def reconcile(self, controls, region_ids, *, complete=True):
        """Only compare the observed scope. Absence is unresolved, never deletion."""
        records = []
        for current in controls:
            record = deepcopy(current);identifier = record["id"]
            previous = self.before.get(identifier)
            if previous:
                protected = previous.get("semantic_source") == "manual"
                for key in ("semantic_name", "intent", "description", "description_source", "description_status",
                            "semantic_source", "aliases", "risk", "actions"):
                    # Feed the planner the prior meaning; it can improve automatic inference later.
                    if key in previous and (protected or key not in {"actions", "risk"}):record[key] = previous[key]
                if previous.get("description_source") == "interaction-observation":
                    for key in ("description", "description_source", "description_status", "description_effect_record"):
                        if key in previous:record[key] = previous[key]
                record = knowledge.merge_control_state(previous, record)
            record["availability"] = dict(state="observed", revision=self.id, at=knowledge.utc_now())
            knowledge.save_control(self.directory, record)
            self.seen.add(identifier);self.missing.discard(identifier);records.append(record)
        if self.mode == "update" and complete:
            for identifier, previous in self.before.items():
                if identifier in self.seen or previous.get("region_id") not in region_ids:continue
                if previous.get("navigation_path") or previous.get("entry_steps"):continue
                if self.control_ids and identifier not in self.control_ids:continue
                self.missing.add(identifier)
                found = knowledge.find_control_record(self.directory, identifier, include_quarantine=True)
                if found:
                    record = found[1]
                    record["availability"] = dict(state="unobserved", revision=self.id,
                        reason="本次对应区域未观察到；可能尚未展开或软件已变化，需要重新定位", at=knowledge.utc_now())
                    knowledge.save_control(self.directory, record)
        self.write()
        return records

    def summarize(self):
        current = {r["id"]: r for r in self.library.read(self.app_id)["controls"]}
        added = self.seen - self.before.keys()
        changed = {i for i in self.seen & self.before.keys() if definition(self.before[i]) != definition(current[i])}
        affected_controls = changed | self.missing
        catalog = workflows.load_catalog(self.directory)
        for identifier, old in self.catalog.items():
            flow = catalog.get(identifier)
            if not flow:continue
            refs=set()
            try:
                refs = {t["control_id"] for _, t in workflows.targets(workflows.expand_calls(flow, catalog))}
                affected = bool(refs & affected_controls) or workflows.definition_hash(old, self.catalog) != workflows.definition_hash(flow, catalog)
                if self.mode in {"update","improve"}:
                    affected = affected or bool(refs & (self.control_ids or self.scope_ids))
            except ValueError:affected = True
            outstanding=workflows.update_validation(flow,list(current.values()),catalog)
            pending_here=outstanding.get("status") in {"pending","failed"} and bool(refs & (self.scope_ids|self.seen|self.control_ids))
            if affected or identifier in self.workflow_ids or pending_here:
                report = workflows.readiness(flow, list(current.values()), catalog)
                previous = flow.get("update_validation", {})
                self.affected[identifier] = dict(status="pending", name=flow["keyword"],
                    reason="需要按新定位运行并校验业务结果" if report["ready"] else "；".join(i["message"] for i in report["issues"]))
                verified = (previous.get("status") == "passed" and previous.get("revision") == self.id and
                            previous.get("definition_hash")==workflows.definition_hash(flow,catalog) and
                            previous.get("control_hash")==workflows.control_hash(flow,list(current.values()),catalog))
                if verified:
                    self.affected[identifier].update(status="passed", reason="实际执行及结果校验通过")
                elif previous.get("revision")==self.id and previous.get("status")!="passed":
                    self.affected[identifier].update(status=previous.get("status","pending"),reason=previous.get("reason",self.affected[identifier]["reason"]))
                else:
                    self.library.save_workflow(self.app_id, {**flow, "update_validation": {
                        **self.affected[identifier], "revision": self.id, "at": knowledge.utc_now()}})
        self.write()
        return dict(revision=self.id, added=sorted(added), updated=sorted(changed), reused=sorted(self.seen - added - changed),
                    unobserved=sorted(self.missing), conflicts=self.conflicts, workflows=self.affected)
