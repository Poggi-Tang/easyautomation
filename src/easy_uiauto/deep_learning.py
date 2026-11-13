"""Region exploration with durable branches and shared, evidence-backed workflows.

The controller has no desktop dependency. Native observation, planning and input
are supplied by region_learning.NativeBackend and can be tested independently.
"""
from __future__ import annotations

from copy import deepcopy
import json
from pathlib import Path

from . import workflows
from . import relearning
from .mcp import image_storage, knowledge


def step_ids(steps, prefix="s"):
    if not isinstance(steps, list):raise ValueError("步骤必须为列表")
    result = deepcopy(steps)
    for number, step in enumerate(result):
        if not isinstance(step, dict):raise ValueError("步骤必须为对象")
        if not isinstance(step.get("action"), str) or not step["action"].strip():
            raise ValueError("步骤缺少 action（动作类型）")
        checks = step.get("checks", [])
        if not isinstance(checks, list):raise ValueError("步骤校验必须为列表")
        for item in [step, *checks]:
            if not isinstance(item, dict):raise ValueError("校验必须为对象")
            if item is not step and (not isinstance(item.get("kind"), str) or not item["kind"].strip()):
                raise ValueError("校验缺少 kind（校验类型）")
            if "target" in item:
                target = item["target"]
                if not isinstance(target, dict) or not isinstance(target.get("control_id", ""), str):
                    raise ValueError("控件引用必须为对象，control_id 必须为字符串")
        step["id"] = f"{prefix}{number}"
        for n, check in enumerate(checks):check["id"] = f"{prefix}{number}.c{n}"
    return result


def normalize_plan(plan):
    """Validate provider data and assign IDs, including non-executable branches."""
    if not isinstance(plan, dict) or not isinstance(plan.get("branches"), list):
        raise ValueError("深度学习规划缺少 branches")
    plan = deepcopy(plan)
    if len(plan["branches"]) > 200:raise ValueError("单个状态的分支超过 200 项")
    if type(plan.get("terminal",False)) is not bool:raise ValueError("terminal 必须为布尔值")
    if "name" in plan and not isinstance(plan["name"], str):raise ValueError("状态名称必须为文本")
    for branch in plan["branches"]:
        if not isinstance(branch, dict) or not isinstance(branch.get("name"), str) or not branch["name"].strip():
            raise ValueError("探索分支需要名称")
        if not isinstance(branch.get("blocked", ""), str):raise ValueError("未完成分支需要说明原因")
        if not isinstance(branch.get("effect", "unknown"), str):raise ValueError("分支影响 effect 必须为文本")
        if not isinstance(branch.get("control_id", ""), str):raise ValueError("分支 control_id 必须为字符串")
        if type(branch.get("replayable",False)) is not bool:raise ValueError("replayable 必须为布尔值")
        branch["steps"] = step_ids(branch.get("steps", []))
    checks = plan.get("completion_checks", [])
    if not isinstance(checks, list):raise ValueError("完成判定必须为校验列表")
    for check in checks:
        if not isinstance(check, dict):raise ValueError("完成校验必须为对象")
        if not isinstance(check.get("target"), dict) or not check["target"].get("control_id"):
            raise ValueError("完成校验缺少 target.control_id（目标控件）")
    if checks:
        plan["completion_checks"] = step_ids([dict(action="inspect", checks=checks)], "result")[0]["checks"]
    for key in ("unresolved", "return_steps", "semantics"):
        if not isinstance(plan.get(key, []), list):raise ValueError(key+" 必须为列表")
    if any(not isinstance(reason, str) for reason in plan.get("unresolved", [])):
        raise ValueError("未完成说明必须为文本")
    plan["return_steps"] = step_ids(plan.get("return_steps", []), "back")
    for semantic in plan.get("semantics", []):
        if not isinstance(semantic, dict) or not isinstance(semantic.get("control_id"), str):
            raise ValueError("语义说明必须为含 control_id 的对象")
        if any(key in semantic and not isinstance(semantic[key], str) for key in ("name", "description")):
            raise ValueError("语义名称和说明必须为文本")
    return plan


def validate_plan(plan, controls):
    """Reject fabricated targets and unsupported actions before any input."""
    plan = normalize_plan(plan)
    if any(not isinstance(r, dict) or not isinstance(r.get("id"), str) or not r["id"] for r in controls):
        raise ValueError("观察到的控件缺少有效 id，请重新扫描区域")
    ids = {r["id"] for r in controls}
    for branch in plan["branches"]:
        if branch.get("blocked"):continue
        flow = dict(keyword=branch["name"], steps=branch["steps"])
        workflows.validate(flow)
        if any(s["action"] in {"call", "manual"} for s in flow["steps"]):raise ValueError("探索只能执行明确的原子动作")
        for _, target in workflows.targets(flow):
            if target.get("control_id") not in ids:raise ValueError("规划引用了当前状态未观察到的控件")
        report = workflows.readiness(flow, controls)
        if not report["ready"]:raise ValueError("探索步骤无效："+"；".join(i["message"] for i in report["issues"]))
    checks = plan.get("completion_checks", [])
    if plan.get("terminal") and not checks:raise ValueError("完成状态必须提供可执行的结果校验")
    for check in checks:
        if not isinstance(check,dict) or not isinstance(check.get("target"),dict) or check["target"].get("control_id") not in ids:
            raise ValueError("结果校验引用了未观察到的控件")
    if checks:
        flow=dict(keyword="结果检查",steps=step_ids([dict(action="inspect",target=checks[0]["target"],checks=checks)]))
        report=workflows.readiness(flow,controls)
        if not report["ready"]:raise ValueError("结果校验定义无效")
    if plan.get("return_steps"):
        flow=dict(keyword="返回",steps=step_ids(plan["return_steps"]))
        report=workflows.readiness(flow,controls)
        if not report["ready"] or any(t.get("control_id") not in ids for _,t in workflows.targets(flow)):
            raise ValueError("返回步骤引用了未观察目标或无效动作")
    return plan


class Explorer:
    def __init__(self, library, app_id, region, backend, *, max_actions=100, max_depth=24,
                 inputs=None, instructions="", resume=True, mode=None, session_id="", control_ids=(),
                 workflow_ids=(), retry_branches=(), progress=lambda _:None):
        if isinstance(max_actions, bool) or not isinstance(max_actions, int) or not 1 <= max_actions <= 2000:
            raise ValueError("探索步数应为 1..2000")
        if isinstance(max_depth, bool) or not isinstance(max_depth, int) or not 1 <= max_depth <= 100:
            raise ValueError("探索深度应为 1..100")
        self.library, self.app_id, self.region, self.backend = library, app_id, region, backend
        self.max_actions, self.max_depth, self.progress = max_actions, max_depth, progress
        self.inputs=deepcopy(inputs or {})
        self.mode = mode or ("resume" if resume else "new")
        if self.mode not in relearning.MODES:raise ValueError("未知学习模式")
        self.selected_controls = set(control_ids)
        self.selected_workflows = set(workflow_ids)
        self.retry_branches = set(retry_branches)
        self.replanned = set();self.visited = set();self.reused_branches = set()
        self.actions = 0;self.stop_reason = "";self.active = set()
        self.directory = library.directory(app_id)
        previous = relearning.select_session(self.directory, region["id"], session_id) if self.mode != "new" else None
        from uuid import uuid4
        identity = previous["id"] if previous else "deep-"+uuid4().hex[:12]
        self.path = self.directory / "learning" / (identity+".json")
        if image_storage._safe_path(self.directory, self.path) is None:raise ValueError("学习记录路径无效")
        self.graph = dict(id=identity, region_id=region["id"], states={}, root="", workflow_ids=[], status="running")
        if previous:
            self.graph = previous
            history = self.directory / "learning" / "history" / (identity+"-"+uuid4().hex[:12]+".json")
            if image_storage._safe_path(self.directory, history) is None:raise ValueError("学习历史路径无效")
            history.parent.mkdir(parents=True, exist_ok=True)
            from .library_tasks import _write
            _write(history, previous)
            for node in self.graph["states"].values():
                for branch in node["branches"]:
                    if branch["status"] == "running":
                        branch.update(status="uncertain", reason="上次任务在操作中停止，请先检查实际结果；不会自动重做")
        known_branches = {b["id"] for n in self.graph["states"].values() for b in n["branches"]}
        if self.retry_branches - known_branches:raise ValueError("待重试分支不属于所选学习记录")
        self.graph.update(mode=self.mode, options=dict(inputs=self.inputs, instructions=instructions,
            **getattr(backend,"policy_identity",{})), scope=dict(control_ids=list(control_ids), workflow_ids=list(workflow_ids)))
        self.graph.setdefault("runs", []).append(dict(mode=self.mode, started_at=knowledge.utc_now(), retry_branches=list(retry_branches)))
        if hasattr(backend,"configure_learning"):
            backend.configure_learning(self.graph, self.mode, control_ids, workflow_ids)
            self.selected_controls = set(backend.revision.control_ids)
        else:
            catalog = workflows.load_catalog(self.directory)
            for identifier in workflow_ids:
                if identifier not in catalog:raise ValueError("所选组合动作已不存在")
                self.selected_controls.update(t["control_id"] for _,t in workflows.targets(workflows.expand_calls(catalog[identifier],catalog)))

    @staticmethod
    def identifier(observation):
        return observation.get("node_id", observation["state_id"])

    def selected(self, node):
        return not self.selected_controls or bool(set(node.get("controls", [])) & self.selected_controls)

    def scope_reachable(self, identifier, trail=()):
        if identifier in trail:return False
        node=self.graph["states"].get(identifier)
        if not node:return False
        return self.selected(node) or any(self.scope_reachable(b.get("target_state"),(*trail,identifier)) for b in node["branches"])

    def branch_in_scope(self, branch):
        if self.mode!="improve" or not self.selected_controls:return True
        ids={t.get("control_id") for _,t in workflows.targets(dict(steps=branch.get("steps",[])))}
        ids.add(branch.get("control_id"))
        return bool(ids & self.selected_controls) or self.scope_reachable(branch.get("target_state"))

    def needs_work(self, identifier, trail=()):
        if identifier in trail:return False
        node=self.graph["states"].get(identifier)
        if not node:return True
        if self.mode in {"improve", "update"} and identifier not in self.visited and self.selected(node):return True
        if node["issues"] or node["terminal"] and not node["verified_terminal"]:return True
        return any(b["status"] not in {"done","cycle"} or b.get("target_state") and
                   self.needs_work(b["target_state"], (*trail,identifier)) for b in node["branches"])

    def persist(self):
        from .library_tasks import _write
        self.graph["updated_at"] = knowledge.utc_now()
        with image_storage._vault_lock(self.library.root):
            self.path.parent.mkdir(parents=True, exist_ok=True)
            _write(self.path, self.graph)
        # Keep a readable checkpoint even if the worker is terminated while
        # waiting for a provider; the finalizer is not guaranteed to run then.
        statuses={"running":"正在学习（若任务已停止，请恢复起点后继续）","partial":"部分完成","complete":"完成"}
        lines=["状态："+statuses.get(self.graph["status"],self.graph["status"]),
               f"已发现 {len(self.graph['states'])} 个状态，归档 {len(self.graph['workflow_ids'])} 个组合动作。",
               "覆盖范围为本次观察到的状态与规划分支，不代表任意输入值的全部组合。",self.stop_reason]
        summary=self.graph.get("changes",{})
        if summary:
            lines.append(f"本次新增 {len(summary['added'])}、更新 {len(summary['updated'])}、复用 {len(summary['reused'])} 个控件；"
                         f"未观察到 {len(summary['unobserved'])} 个，冲突 {len(summary['conflicts'])} 项。")
            lines.extend(f"- 流程 {v['name']}：{v['status']}；{v['reason']}" for v in summary["workflows"].values())
            if summary["unobserved"]:lines.append("本次未观察到，保留待定位："+"、".join(summary["unobserved"]))
            lines.extend(f"- 合并冲突 {c['id']}：{c['reason']}" for c in summary["conflicts"])
            for conflict in summary["conflicts"]:
                lines.extend(f"  建议步骤 {n}：{s.get('description') or s.get('action')} / "
                             f"{s.get('target',{}).get('label') or s.get('target',{}).get('control_id') or s.get('workflow_id','')}"
                             for n,s in enumerate(conflict.get("proposed_steps",[]),1))
        for node in self.graph["states"].values():
            for branch in node["branches"]:
                if branch["status"] not in {"done","cycle"}:
                    lines.append(f"- {node['name']} / {branch['name']}：{branch.get('reason') or branch['status']}")
            lines.extend(f"- {node['name']}：{reason}" for reason in node["issues"])
        self.library.save_usage(self.app_id,self.region.get("name","区域")+" · 深度学习进度","\n\n".join(lines),
                                document_id=self.graph["id"],workflow_ids=self.graph["workflow_ids"])

    def node(self, observation, route):
        identifier = self.identifier(observation)
        previous = self.graph["states"].get(identifier)
        if identifier in self.replanned:return previous
        if previous and self.mode == "improve" and not self.selected(previous):return previous
        if previous and self.mode == "resume" and not self.needs_work(identifier):return previous
        self.progress("分析当前状态及可交互分支")
        observation["previous_learning"] = deepcopy(previous or {})
        plan = validate_plan(self.backend.plan(observation, route), observation["controls"])
        branches = [];remaining = list((previous or {}).get("branches", []))
        for item in plan["branches"]:
            old = relearning.branch_match(item, remaining)
            if old:remaining.remove(old)
            key = old["id"] if old else knowledge.stable_id("branch", relearning.steps_key(item.get("steps", [])), item["name"])
            if any(b["id"] == key for b in branches):continue
            branch = {**item, "id":key, "status":"blocked" if item.get("blocked") else "pending", "reason":item.get("blocked", "")}
            targeted = not self.selected_controls or bool({t.get("control_id") for _,t in workflows.targets(
                dict(steps=item.get("steps", [])))} & self.selected_controls)
            if self.mode=="improve" and not targeted and not old:continue
            if self.mode == "improve" and not targeted and old:
                branch = deepcopy(old)
            elif old:
                same = relearning.steps_key(old.get("steps", [])) == relearning.steps_key(item.get("steps", []))
                if branch.get("blocked"):
                    if old.get("evidence"):branch["evidence"]=old["evidence"]
                elif key in self.retry_branches:
                    branch["retry_confirmed_at"] = knowledge.utc_now()
                elif old["status"] in {"uncertain", "running"} or old["status"] == "observed" and not old.get("replayable"):
                    branch = {**old, "status":"uncertain", "reason":"上次操作可能已生效；先核对结果，或在该结果界面继续。确认未生效后才可重试"}
                elif old["status"] in {"done","cycle","observed"}:
                    refresh = self.mode in {"improve", "update"} and targeted
                    if same and not refresh:
                        branch=deepcopy(old);self.reused_branches.add(key)
                    elif old.get("replayable"):
                        branch.update(target_state=old.get("target_state"))
                    else:
                        branch={**old,"status":"uncertain","reason":"该分支包含已执行的业务操作；重新验证前请确认测试数据已恢复"}
            branches.append(branch)
        for old in remaining:
            # A provider omitting a known entry is not evidence that it disappeared.
            branches.append({**old, "status":"uncertain" if old["status"]=="uncertain" else "blocked",
                             "reason":"本次未规划旧入口，请补充说明或确认该入口是否仍存在"})
        node = dict(id=identifier, name=plan.get("name") or observation.get("name") or "界面状态",
                    controls=[r["id"] for r in observation["controls"]], branches=branches,
                    return_steps=step_ids(plan.get("return_steps", []), "back"),
                    completion_checks=plan.get("completion_checks", []), terminal=bool(plan.get("terminal")),
                    verified_terminal=False, issues=list(observation.get("issues", []))+list(plan.get("unresolved", [])),
                    route=deepcopy(route), region_ids=observation.get("region_ids", []), fingerprint=observation["state_id"])
        if previous and self.mode == "resume" and previous.get("verified_terminal") and previous.get("completion_checks")==node["completion_checks"]:
            node.update(verified_terminal=True, terminal_evidence=previous.get("terminal_evidence"))
        # Every observed interactive control must be accounted for, even when it
        # requires inputs or an unsupported gesture. Omission is not coverage.
        covered = {t.get("control_id") for b in branches for _,t in workflows.targets(dict(steps=b.get("steps", [])))}
        covered.update(b.get("control_id") for b in branches if b.get("blocked"))
        if not node["terminal"]:
            for control in observation["controls"]:
                if self.mode=="improve" and self.selected_controls and control["id"] not in self.selected_controls:continue
                if control.get("actions") and control.get("is_enabled_at_scan") is not False and control["id"] not in covered:
                    node["issues"].append("未规划交互："+(control.get("semantic_name") or control["id"]))
            if not branches:node["issues"].append("尚无可执行分支，也未确认业务结果")
        self.graph["states"][identifier] = node;self.replanned.add(identifier);self.persist()
        return node

    def visit(self, observation, route, depth=0):
        node = self.node(observation, route);identifier = node["id"]
        self.visited.add(identifier)
        if identifier in self.active:return
        if node["terminal"] and not node["verified_terminal"]:
            node["terminal_evidence"]=self.backend.verify(observation, node["completion_checks"])
            node["verified_terminal"] = True;node["verified_run"]=len(self.graph["runs"]);self.persist()
            self.archive();self.persist()
        if depth >= self.max_depth and any(b["status"]=="pending" for b in node["branches"]):
            self.stop_reason = "达到本次探索深度上限";return
        self.active.add(identifier)
        try:
            for branch in node["branches"]:
                if self.stop_reason:return
                if not self.branch_in_scope(branch):continue
                if branch["status"] in {"blocked", "failed", "uncertain", "cycle"}:continue
                if branch["status"] == "done":
                    if not branch.get("target_state") or not self.needs_work(branch["target_state"]):continue
                    if not branch.get("replayable"):
                        branch.update(status="uncertain", reason="后续尚未完成，不能重复提交；请打开上次结果界面继续")
                        self.persist();continue
                count = len(branch.get("steps", []))
                if self.actions+count > self.max_actions:
                    self.stop_reason = "达到本次探索步数上限";return
                self.backend.assert_state(observation)
                self.progress(f"深度学习 {self.actions}/{self.max_actions}：{branch['name']}")
                branch["status"] = "running";self.persist()
                self.actions += count
                try:
                    evidence = self.backend.execute(branch["steps"])
                    child_route = [*route, *deepcopy(branch["steps"])]
                    child = self.backend.observe(route=child_route)
                    branch.update(target_state=self.identifier(child), evidence=evidence, status="observed",executed_run=len(self.graph["runs"]))
                    self.persist()
                    if self.identifier(child) in self.active:
                        branch.update(status="cycle", reason="回到已在探索的状态，已记录交互，未重复进入")
                    else:
                        self.visit(child, child_route, depth+1)
                        branch["status"] = "done" if not self.stop_reason else "observed"
                    self.persist()
                    if self.stop_reason:return
                    child_node = self.graph["states"].get(self.identifier(child), {})
                    self.backend.restore(observation, child, child_node.get("return_steps", []), route)
                except InterruptedError as error:
                    if branch["status"]=="running":
                        evidence=getattr(error,"execution_result",{})
                        branch.update(status="failed" if evidence.get("input_attempted") is False else "uncertain",
                                      reason="操作期间停止，请检查当前结果",evidence=evidence)
                    self.persist()
                    raise
                except Exception as error:
                    if branch["status"] == "running":
                        execution=getattr(error,"execution_result",{})
                        branch.update(status="failed" if execution.get("input_attempted") is False else "uncertain",
                                      reason=str(error), evidence=execution)
                    else:branch["reason"] = str(error)
                    self.stop_reason = str(error);self.persist();return
        finally:self.active.discard(identifier)

    def archive(self):
        """Each suffix is one workflow; all converging entry paths call that ID."""
        memo = {};created = [];catalog=workflows.load_catalog(self.directory)
        def suffixes(identifier, trail):
            if identifier in memo:return memo[identifier]
            if identifier in trail:return []
            node = self.graph["states"][identifier];result=[]
            if node.get("verified_terminal"):
                checks = deepcopy(node["completion_checks"])
                steps = [dict(action="inspect", target=checks[0]["target"], checks=checks, description="确认业务结果")]
                result.append(save(identifier, "result", node["name"], steps, verified=node.get("verified_run")==len(self.graph["runs"])))
            for branch in node["branches"]:
                target = branch.get("target_state")
                if branch["status"] not in {"done", "observed"} or target not in self.graph["states"]:continue
                for child in suffixes(target, [*trail, identifier]):
                    verified=(branch.get("executed_run")==len(self.graph["runs"]) and
                              catalog[child].get("learning_evidence",{}).get("verified_run")==len(self.graph["runs"]))
                    result.append(save(identifier, branch["id"]+child, branch["name"], [*branch["steps"],
                        dict(action="call", workflow_id=child, description="复用后续流程", parameters={})],
                        verified=verified,branch_id=branch["id"]))
                    if len(created) >= 500:
                        self.stop_reason = "组合动作归档达到 500 项上限";return result
            memo[identifier]=result;return result
        def save(state, key, name, steps, verified=False,branch_id=""):
            identifier=knowledge.stable_id("deep-flow", self.graph["id"], state, key)
            if len(created)>=500:raise ValueError("组合动作归档达到 500 项上限")
            existing=catalog.get(identifier)
            if not existing and branch_id:
                candidates=[f for f in catalog.values() if f.get("learning_session")==self.graph["id"] and
                            f.get("entry_state")==state and (f.get("learning_branch_id")==branch_id or
                            not f.get("learning_branch_id") and relearning.steps_key(f.get("steps",[])[:-1],anchor=True)==
                            relearning.steps_key(steps[:-1],anchor=True))]
                if len(candidates)==1 and candidates[0]["id"] not in created:
                    existing=candidates[0];identifier=existing["id"]
            if existing and ((existing.get("learning_authored_hash") and existing["learning_authored_hash"] != relearning.authored_key(existing)) or
                             not existing.get("learning_authored_hash") and existing.get("learning_definition_hash") != existing.get("definition_hash")):
                # A user-edited generated workflow belongs to the user now.
                if hasattr(self.backend,"revision"):
                    conflict=dict(kind="workflow",id=identifier,reason="保留人工修改的组合动作",proposed_steps=step_ids(steps))
                    if conflict not in self.backend.revision.conflicts:self.backend.revision.conflicts.append(conflict)
                return identifier
            steps=step_ids(steps)
            for step in steps:
                for item,field in [(step,"text"),*((check,"expected") for check in step.get("checks",[]))]:
                    value=item.get(field)
                    if isinstance(value,str):
                        for parameter,example in self.inputs.items():
                            if isinstance(example,str) and example and value==example:item[field]="${"+parameter+"}";break
            flow=dict(**{k:v for k,v in (existing or {}).items() if k not in {"id","keyword","name","steps","precondition","parameters","source","learning_session","entry_state","learning_verified"}},
                      id=identifier, keyword=(existing or {}).get("keyword") or name+" · "+identifier[-8:], name=name, steps=steps,
                      precondition="软件已启动", parameters=self.inputs, source="deep-learning",
                      learning_session=self.graph["id"], entry_state=state, learning_verified=True)
            flow["learning_definition_hash"]=workflows.definition_hash(flow,catalog)
            flow["learning_steps_hash"]=relearning.steps_key(steps)
            flow["learning_branch_id"]=branch_id
            flow["learning_authored_hash"]=relearning.authored_key(flow)
            flow["learning_evidence"] = dict(state=state, at=knowledge.utc_now(), verified_run=len(self.graph["runs"]) if verified else None)
            if verified and self.mode=="update" and getattr(self.backend,"revision",None):
                flow["update_validation"]=dict(status="passed",revision=self.backend.revision.id,
                    reason="本次学习已执行完整路径并通过结果校验",definition_hash=workflows.definition_hash(flow,catalog),
                    control_hash=workflows.control_hash(flow,self.library.read(self.app_id)["controls"],catalog),at=knowledge.utc_now())
            saved=self.library.save_workflow(self.app_id, flow)
            catalog[identifier]=saved;created.append(identifier)
            if identifier not in self.graph["workflow_ids"]:self.graph["workflow_ids"].append(identifier)
            return identifier
        if self.graph["root"] in self.graph["states"]:
            roots=suffixes(self.graph["root"], [])
            entries=getattr(self.backend,"root_entry",[])
            if entries:
                roots=[save(self.graph["root"],"entry-"+flow,self.region.get("name","区域")+"入口",
                       [*entries,dict(action="call",workflow_id=flow,description="执行已学习流程",parameters={})]) for flow in roots]
            self.graph["entry_workflow_ids"]=roots
        self.graph["workflow_ids"] = sorted(set(self.graph.get("workflow_ids", [])) | set(created))

    def run(self):
        error=None
        try:
            observation = self.backend.observe(root=True, route=[])
            identifier = self.identifier(observation)
            if self.graph["root"] and identifier != self.graph["root"] and identifier not in self.graph["states"]:
                raise ValueError("当前界面不在已有学习记录中；请恢复已知状态，软件已更新时选择“适配软件更新”")
            self.graph["root"] = self.graph["root"] or identifier;self.graph["status"]="running";self.persist()
            route=self.graph["states"].get(identifier,{}).get("route",[])
            self.visit(observation, route)
            # Continue directly at a saved result state without repeating its commit.
            if identifier != self.graph["root"] and not self.stop_reason:
                for node in self.graph["states"].values():
                    for branch in node["branches"]:
                        if branch.get("target_state")==identifier and not self.needs_work(identifier):
                            branch.update(status="done",reason="从已观察到的结果状态继续完成")
                if self.needs_work(self.graph["root"]):
                    self.stop_reason="当前状态的后续学习已保存；请恢复原区域起点，再继续其余入口"
        except BaseException as failure:
            error=failure;self.stop_reason=str(failure)
        finally:
            try:self.archive()
            except Exception as failure:
                self.stop_reason="；".join(s for s in (self.stop_reason,"组合动作归档未完成："+str(failure)) if s)
            if hasattr(self.backend,"revision"):
                try:
                    self.graph["changes"] = self.backend.revision.summarize()
                    if self.graph["changes"]["workflows"] and not self.stop_reason:
                        self.backend.validate_affected_workflows(self.graph, self.max_actions-self.actions)
                        self.graph["changes"] = self.backend.revision.summarize()
                except Exception as failure:
                    self.stop_reason="；".join(s for s in (self.stop_reason,"知识变化检查未完成："+str(failure)) if s)
            unfinished=[dict(state=n["name"], branch=b["name"], branch_id=b["id"], state_id=n["id"],status=b["status"], reason=b.get("reason", ""))
                        for n in self.graph["states"].values() for b in n["branches"]
                        if b["status"] not in {"done", "cycle"} and self.branch_in_scope(b)]
            issues=[dict(state=n["name"],reason=str(reason)) for n in self.graph["states"].values() for reason in n["issues"]
                    if self.mode!="improve" or self.selected(n)]
            if self.mode=="improve" and self.selected_controls:
                observed={i for key in self.visited for i in self.graph["states"][key].get("controls",[])}
                missing=self.selected_controls-observed
                if missing:issues.append(dict(state="指定内容",reason="以下控件尚未到达或已变化，请打开对应状态后继续："+"、".join(sorted(missing))))
            for flow in self.graph.get("changes",{}).get("workflows",{}).values():
                if flow["status"]!="passed":issues.append(dict(state=flow["name"],reason=flow["reason"]))
            if self.graph.get("changes",{}).get("unobserved"):
                issues.append(dict(state="更新适配",reason="仍有旧控件未重新定位，请打开对应区域后继续"))
            if self.graph.get("changes",{}).get("conflicts"):issues.append(dict(state="知识合并",reason="存在保留的人工修改或定位冲突，请查看学习报告"))
            self.graph.update(status="partial" if unfinished or issues or self.stop_reason else "complete",
                              unfinished=unfinished, issues=issues, stop_reason=self.stop_reason)
            self.persist()
        if error:raise error
        change_message=""
        changes=self.graph.get("changes",{})
        if changes:
            change_message=(f" 新增 {len(changes['added'])}、更新 {len(changes['updated'])}、复用 {len(changes['reused'])} 个控件；"
                f"{len(changes['unobserved'])} 个待定位，{sum(f['status']!='passed' for f in changes['workflows'].values())} 个流程待验证。")
        return dict(ok=True, complete=self.graph["status"]=="complete", states=len(self.graph["states"]),
                    actions=self.actions, workflow_ids=self.graph["workflow_ids"], unfinished=unfinished, issues=issues,
                    report_path=str(self.path), session_id=self.graph["id"],changes=self.graph.get("changes",{}),
                    message=f"{relearning.MODES[self.mode]}{'完成' if self.graph['status']=='complete' else '部分完成'}；"
                    f"发现 {len(self.graph['states'])} 个状态，{len(unfinished)} 条分支未完成。"+(self.stop_reason or "")+change_message)
