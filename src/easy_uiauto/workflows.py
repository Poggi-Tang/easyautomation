"""Test-case keywords: explicit actions, knowledge bindings and UI assertions."""

from __future__ import annotations

from copy import deepcopy
import hashlib
import json
import re
import time
from uuid import uuid4

from .mcp import knowledge

ACTIONS = {"click", "double-click", "right-click", "set-text", "press-key", "hotkey", "inspect", "manual"}
CHECKS = {"exists", "not-visible", "text", "window-title", "read-only", "menu-items",
          "node-exists", "node-absent", "selected", "manual"}


def click_step(record, *, action="click"):
    return dict(id="step-"+uuid4().hex[:10], action=action,
                description=record.get("semantic_name") or record.get("name") or record["id"],
                target=dict(control_id=record["id"], label=record.get("name", "")), checks=[], timeout_seconds=10)


def substitute(value, parameters):
    """Expand strings in input and expectations before any input is attempted."""
    if isinstance(value, str):
        def replace(match):
            if match[1] not in parameters: raise ValueError("缺少流程参数："+match[1])
            return str(parameters[match[1]])
        return re.sub(r"\$\{([\w.]+)\}", replace, value)
    if isinstance(value, list): return [substitute(v, parameters) for v in value]
    if isinstance(value, dict): return {k:substitute(v, parameters) for k,v in value.items()}
    return value


def _cell(value):
    return re.sub(r"<br\s*/?>", "\n", value, flags=re.I).strip()


def _numbered(value):
    return [s.strip() for s in re.split(r"(?:^|\n)\s*\d+[.．、]\s*", _cell(value)) if s.strip()]


def _target(label, kind="", **extra):
    return {"label": label, "control_type": kind, "control_id": "", **extra}


def import_case(markdown: str, app_id: str, keyword="") -> dict:
    """Compile the documented Chinese table format; unsupported prose stays manual."""
    tables, headers = {}, None
    for line in markdown.splitlines():
        if not line.strip().startswith("|"): continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if all(re.fullmatch(r":?-+:?", c.replace(" ", "")) for c in cells): continue
        if "用例名称" in cells or "动作标题" in cells or "步骤" in cells and "预期" in cells:
            headers = cells; continue
        if headers:
            for k, v in zip(headers, cells): tables[k] = _cell(v)
            headers = None
    title = tables.get("动作标题", tables.get("用例名称", "")).strip()
    actions, expectations = _numbered(tables.get("步骤", "")), _numbered(tables.get("预期", ""))
    if not title or not actions or len(actions) != len(expectations):
        raise ValueError("请粘贴包含用例名称、前置条件、步骤、预期的 Markdown 表格，步骤和预期需逐条对应")
    steps = []
    for index, (description, expected) in enumerate(zip(actions, expectations), 1):
        labels = re.findall(r"【([^】]+)】", description)
        label = labels[0] if labels else ""
        action = "right-click" if re.search(r"右键|右击", description) else "click" if "点击" in description else "inspect" if "查看" in description else "manual"
        target = _target(label, "MenuItemControl" if "菜单项" in description else "")
        if action == "inspect" and "对话框" in description:
            target["control_type"] = "WindowControl"
        if "空白" in description: target["hit"] = "blank"
        checks = []
        if "右键菜单" in expected and re.search(r"含[：:]", expected):
            menu = re.split(r"含[：:]", expected, maxsplit=1)[1]
            items = [s.strip().strip("【】") for s in re.split(r"[、，,]", menu) if s.strip()]
            checks.append(dict(kind="menu-items", target=_target(f"{label}右键菜单", "MenuControl"), expected=items))
        elif "标题为" in expected:
            match = re.search(r"标题为[「“\"]([^」”\"]+)[」”\"]", expected)
            if match:
                checks.append(dict(kind="window-title", target=_target(match[1], "WindowControl"), expected=match[1]))
            else:
                checks.append(dict(kind="manual", description=expected))
        elif re.match(r"含[：:]", expected):
            for item in re.split(r"[、]", re.split(r"[：:]", expected, maxsplit=1)[1]):
                name = re.sub(r"[（(].*?[）)]", "", item).strip().strip("【】")
                if not name: continue
                if "只读" in item or "可编辑" in item:
                    checks.append(dict(kind="read-only", target=_target(name, "EditControl"), expected="只读" in item))
                else:
                    checks.append(dict(kind="exists", target=_target(name, "ButtonControl" if "【" in item else "")))
        else:
            checks.append(dict(kind="manual", description=expected))
        for n, check in enumerate(checks, 1): check["id"] = f"step-{index}.check-{n}"
        steps.append(dict(id=f"step-{index}", action=action, target=target, description=description,
                          expected_description=expected, checks=checks, timeout_seconds=10))
    default_keyword = re.sub(r"^【([^】]+)】", r"\1.", title)
    return dict(id=knowledge.stable_id("workflow", keyword or default_keyword), kind="workflow", schema_version=1,
                app_id=app_id, keyword=keyword or default_keyword, name=title,
                precondition=tables.get("前置条件", ""), parameters={}, steps=steps,
                source_case=markdown, status="draft")


def targets(flow):
    for step in flow.get("steps", []):
        if isinstance(step.get("target"), dict): yield step["id"], step["target"]
        for check in step.get("checks", []):
            if isinstance(check.get("target"), dict): yield check["id"], check["target"]


def suggest_bindings(flow, controls):
    """Bind only one exact knowledge label; never infer a role from a screenshot."""
    flow = deepcopy(flow)
    for _identifier, target in targets(flow):
        if target.get("control_id") or not target.get("label", "").strip(): continue
        matches = [r for r in controls if target["label"] in {r.get("name"), r.get("semantic_name"), *r.get("aliases", [])}
                   and (not target.get("control_type") or target["control_type"] == r.get("control_type"))
                   and not knowledge.runtime_quarantined(r)]
        if len(matches) == 1: target["control_id"] = matches[0]["id"]
    return flow


def validate(flow):
    if not isinstance(flow, dict) or not isinstance(flow.get("keyword"), str) or not flow["keyword"].strip():
        raise ValueError("流程关键字不能为空")
    if not isinstance(flow.get("parameters", {}), dict): raise ValueError("流程参数必须为对象")
    if not isinstance(flow.get("precondition", ""), str): raise ValueError("前置条件必须为文本")
    steps = flow.get("steps")
    if not isinstance(steps, list) or not 1 <= len(steps) <= 200: raise ValueError("流程需要 1 至 200 个步骤")
    seen = set()
    for step in steps:
        if not isinstance(step, dict): raise ValueError("步骤必须为对象")
        if not isinstance(step.get("action"), str) or step["action"] not in ACTIONS: raise ValueError("未知流程动作")
        if not isinstance(step.get("checks", []), list) or len(step.get("checks", [])) > 100: raise ValueError("校验项必须为列表且不超过 100 项")
        for item in [step, *step.get("checks", [])]:
            if not isinstance(item, dict): raise ValueError("校验必须为对象")
            identifier = item.get("id")
            if not isinstance(identifier, str) or not identifier or identifier in seen: raise ValueError("步骤和校验 ID 必须唯一")
            seen.add(identifier)
            if item is not step and (not isinstance(item.get("kind"), str) or item["kind"] not in CHECKS): raise ValueError("未知校验类型")
            target = item.get("target")
            if target is not None and (not isinstance(target, dict) or not isinstance(target.get("control_id", ""), str)):
                raise ValueError("控件引用必须为对象，control_id 必须为字符串")
            if target and (not isinstance(target.get("hit", "center"), str) or target.get("hit", "center") not in {"center", "blank"}): raise ValueError("未知点击位置模式")
            if target and not isinstance(target.get("label", ""), str): raise ValueError("控件标签必须为文本")
        timeout = step.get("timeout_seconds", 10)
        if isinstance(timeout, bool) or not isinstance(timeout, (int, float)) or not 0 <= timeout <= 60: raise ValueError("步骤等待时间必须在 0 至 60 秒之间")
    return flow


def definition_hash(flow):
    """Keep evidence tied to the exact steps and bindings that were executed."""
    definition = {key: flow.get(key) for key in ("keyword", "precondition", "parameters", "steps")}
    return hashlib.sha256(json.dumps(definition, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def control_hash(flow, controls):
    """Explicit steps reference control IDs; locator changes invalidate old evidence."""
    records = {r["id"]: r for r in controls}
    referenced = {}
    for _identifier, target in targets(flow):
        identifier = target.get("control_id", "")
        record = records.get(identifier, {})
        keys = ("location", "window_scope", "control_type", "status")
        referenced[identifier] = {k: record.get(k) for k in keys}
        if not record.get("location"):
            referenced[identifier].update({k: record.get(k) for k in ("image", "visual_target", "normalized_rect")})
    return hashlib.sha256(json.dumps(referenced, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def readiness(flow, controls):
    from .library_actions import validate_plan

    validate(flow)
    issues = []
    for identifier, target in targets(flow):
        control_id = target.get("control_id")
        if not control_id:
            issues.append(dict(id=identifier, message="待绑定控件："+target.get("label", ""))); continue
        try:
            record = validate_plan(controls, control_id, path=[])[0]
            if target.get("control_type") and target["control_type"] != record.get("control_type"):
                raise ValueError("绑定的控件类型与用例要求不一致")
        except (ValueError, RuntimeError) as error:
            issues.append(dict(id=identifier, message=str(error)))
    for step in flow["steps"]:
        for item in [step, *step.get("checks", [])]:
            if item.get("action") == "manual" or item.get("kind") == "manual":
                issues.append(dict(id=item["id"], message="该描述尚未转换为可执行动作或校验"))
            elif not item.get("target"):
                issues.append(dict(id=item["id"], message="缺少目标控件"))
            if item.get("kind") == "read-only" and not isinstance(item.get("expected"), bool):
                issues.append(dict(id=item["id"], message="只读预期必须为 true 或 false"))
            if item.get("kind") in {"text", "window-title"} and not isinstance(item.get("expected"), str):
                issues.append(dict(id=item["id"], message="文本预期必须为字符串"))
            if item.get("kind") in {"node-exists", "node-absent"} and (
                    not isinstance(item.get("expected"), str) or not item["expected"].strip()):
                issues.append(dict(id=item["id"], message="树节点预期必须为非空名称"))
            if item.get("kind") == "menu-items" and (not isinstance(item.get("expected"), list) or
                    not item["expected"] or any(not isinstance(v, str) or not v for v in item["expected"])):
                issues.append(dict(id=item["id"], message="菜单预期必须为非空文本列表"))
    precondition = flow.get("precondition", "").strip()
    if precondition not in {"", "软件已启动"}:
        issues.append(dict(id="precondition", message="前置条件需要明确转换为校验："+precondition))
    return dict(ready=not issues, issues=issues, targets=len(list(targets(flow))))


def _names(control, maximum=500):
    result, pending, count = [], [control], 0
    while pending and count < maximum:
        current = pending.pop(); count += 1
        if bool(getattr(current, "IsOffscreen", False)): continue
        name = re.sub(r"&", "", str(current.Name or "").split("\t")[0]).strip()
        if name and getattr(current, "ControlTypeName", "") == "MenuItemControl": result.append(name)
        pending.extend(reversed(current.GetChildren()))
    if pending: raise ValueError("菜单控件超过读取上限，不能确认预期")
    return result


def _check(check, control):
    from .mcp.ui_cli import _read_control_text
    from .locator_images import ImageTarget, target_evidence
    from .mcp.results import ActionError

    kind, expected = check["kind"], check.get("expected")
    if isinstance(control,ImageTarget):
        if kind=="exists":return dict(ok=True,exists=True,**target_evidence(control))
        if kind=="not-visible":return dict(ok=False,visible=True,**target_evidence(control))
        raise ActionError("operation_blocked","图片匹配只能确认可见目标，无法读取此 UIA 属性预期",reason="image_property_unavailable")
    if kind == "exists": return {"ok": True, "exists": True}
    if kind == "not-visible":
        visible = not bool(getattr(control, "IsOffscreen", False))
        return dict(ok=not visible, visible=visible)
    if kind == "selected":
        pattern = control.GetSelectionItemPattern()
        if pattern is None: raise ValueError("控件不提供 SelectionItemPattern，无法确认选中状态")
        return dict(ok=bool(pattern.IsSelected), selected=bool(pattern.IsSelected))
    if kind in {"node-exists", "node-absent"}:
        pending, count, found = [(control, 0)], 0, []
        deadline = time.monotonic()+2
        while pending:
            current, depth = pending.pop(); count += 1
            if count > 1000 or depth > 32 or time.monotonic() > deadline:
                raise ValueError("树节点读取达到上限，不能判定节点不存在")
            if current is not control and str(current.Name or "") == expected:
                found.append(str(current.Name))
                if kind == "node-exists": break
            pending.extend((child, depth+1) for child in reversed(current.GetChildren()))
        return dict(ok=bool(found) if kind == "node-exists" else not found,
                    expected=expected, observed=found, exists=bool(found), inspected_nodes=count)
    if kind == "window-title":
        title = str(control.Name or "")
        return dict(ok=title == expected, observed=title, expected=expected)
    if kind == "menu-items":
        names = _names(control)
        missing = [name for name in expected if name not in names]
        return dict(ok=not missing, observed=names, missing=missing)
    if kind == "read-only":
        pattern = control.GetValuePattern()
        if pattern is None: raise ValueError("控件不提供 ValuePattern，无法确认只读或可编辑")
        read_only = bool(pattern.IsReadOnly)
        enabled = bool(control.IsEnabled)
        return dict(ok=read_only == expected and (expected or enabled), read_only=read_only, enabled=enabled)
    readable, text = _read_control_text(control)
    if not readable: text = str(control.Name or "")
    return dict(ok=text == expected, observed=text, expected=expected)


def _blank_offset(control):
    import uiautomation
    from .mcp.scanner import _rect

    r = _rect(control)
    for fy in (.85, .7, .5, .3):
        for fx in (.5, .25, .75):
            x = max(1, min(r["width"]-1, round(r["width"]*fx)))
            y = max(1, min(r["height"]-1, round(r["height"]*fy)))
            hit = uiautomation.ControlFromPoint(r["left"]+x, r["top"]+y)
            if hit is not None and uiautomation.ControlsAreSame(control, hit): return x, y
    raise ValueError("未找到可确认的容器空白处，未发送右键")


def execute(flow, controls, expected_window, *, parameters=None, progress=None, checkpoint=None,
            before_action=None, after_action=None, after_check=None, skip_open_entries=False, directory=None):
    """Run explicit steps once through Controller; only observations are polled."""
    from .library_actions import bind_window, owned_windows, resolve_record
    from .ctrl import Controller
    from .locator_images import ImageTarget, target_arguments, target_evidence
    from .mcp.results import ActionError
    from _ctypes import COMError
    import uiautomation

    report = readiness(flow, controls)
    if not report["ready"]: raise ValueError("流程尚不可运行："+"；".join(i["message"] for i in report["issues"]))
    progress, checkpoint = progress or (lambda _m: None), checkpoint or (lambda _r: None)
    if parameters is not None and not isinstance(parameters, dict): raise ValueError("流程参数必须为对象")
    parameters = {**flow.get("parameters", {}), **(parameters or {})}
    flow = deepcopy(flow)
    original_hash = definition_hash(flow)
    flow["steps"] = substitute(flow["steps"], parameters)
    texts = {step["id"]: str(step.get("text", "")) for step in flow["steps"]}
    records = {r["id"]: r for r in controls}
    result = dict(ok=False, run_id=uuid4().hex, keyword=flow["keyword"], definition_hash=original_hash,
                  control_hash=control_hash(flow, controls),
                  parameters=parameters, steps=[], input_attempted=False)
    main, activated = None, False

    resolve_options={"directory":directory} if directory is not None else {}
    def resolve(target): return resolve_record(records[target["control_id"]], expected_window, main, **resolve_options)

    def observe(item, timeout, assertion=False):
        deadline = time.monotonic()+timeout
        while True:
            progress("正在检查："+item.get("description", item["target"].get("label", item["id"])))
            try:
                if assertion and item.get("kind") == "not-visible":
                    control = resolve_record(records[item["target"]["control_id"]], expected_window, main, require_visible=False, **resolve_options)
                else:
                    control = resolve(item["target"])
                evidence = _check(item, control) if assertion else {"ok": True}
                if evidence["ok"]: return control, evidence
            except ActionError as error:
                if error.code != "control_not_found" or error.reason: raise
                if assertion and item.get("kind") == "not-visible":
                    return None, dict(ok=True, visible=False, code=error.code)
                evidence = dict(ok=False, code=error.code, error=str(error))
            except COMError as error:
                # A disappearing Qt dialog can leave a stale provider between
                # native enumeration and UIA binding. Re-observe only; a stale
                # COM object alone is never proof that an assertion passed.
                if error.hresult & 0xffffffff not in {0x80040201,0x80010108,0x800401FD}:
                    raise
                evidence = dict(ok=False, code="stale_observation", error=str(error))
            if time.monotonic() >= deadline:
                error = ValueError("目标未出现或预期未满足："+json.dumps(evidence, ensure_ascii=False))
                error.assertion_evidence = evidence
                raise error
            time.sleep(.1)

    try:
        main = bind_window(expected_window)
        result["precondition"] = {"ok": True, "condition": "软件已启动"}
        for index,step in enumerate(flow["steps"]):
            item = {"id": step["id"], "description": step.get("description", ""), "action": step["action"],
                    "input_attempted": False, "checks": [], "ok": False}
            result["steps"].append(item); checkpoint(result)
            if skip_open_entries and step.get("skip_when_next_visible") and index+1<len(flow["steps"]):
                try:next_control=resolve_record(records[flow["steps"][index+1]["target"]["control_id"]], expected_window, main)
                except ActionError as error:
                    if error.code!="control_not_found" or error.reason:raise
                else:
                    if bool(getattr(next_control,"IsEnabled",True)):
                        item.update(ok=True,skipped=True,reason="next-control-already-visible")
                        checkpoint(result);continue
            if not activated and uiautomation.IsIconic(expected_window["handle"]):
                control = resolve_record(records[step["target"]["control_id"]], expected_window, main, require_visible=False, **resolve_options)
            else:
                control, _evidence = observe(step, step.get("timeout_seconds", 10))
            action = step["action"]
            text = texts[step["id"]]
            if not activated and (action!="inspect" or after_action or after_check):
                scope=records[step["target"]["control_id"]].get("window_scope",{"kind":"main"})
                if scope.get("kind")=="owned":
                    from .region_actions import target_window
                    target=target_window(expected_window,scope)
                    Controller.activate_window_handle(target["handle"])
                else:
                    if expected_window.get("process_id") and expected_window.get("class_name"):
                        from .library_actions import activate_application
                        activate_application(expected_window)
                    else:Controller.activate_window_handle(expected_window["handle"], control=main)
                activated = True
                control, _evidence = observe(step, step.get("timeout_seconds", 10))
            item.update(target_evidence(control))
            if action != "inspect":
                allowed = {expected_window["handle"], *(w["handle"] for w in owned_windows(expected_window))}
                if int(uiautomation.GetForegroundWindow() or 0) not in allowed:
                    raise ActionError("window_invalid", "前台已切换到其他软件，停止流程")
                record = records[step["target"]["control_id"]]
                if before_action: item["before"] = before_action(record, control, main)
                options = {}
                if step["target"].get("hit") == "blank":
                    if isinstance(control,ImageTarget):
                        raise ActionError("operation_blocked","图片匹配无法确认容器空白处，未发送操作")
                    options["x"], options["y"] = _blank_offset(control)
                if isinstance(control,ImageTarget) and action=="set-text":
                    raise ActionError("operation_blocked","图片目标不提供 ValuePattern；请先修复输入框的 UIA 定位")
                if int(uiautomation.GetForegroundWindow() or 0) not in allowed:
                    raise ActionError("window_invalid", "观察期间前台已切换，未发送当前操作")
                from .popup_guard import assert_unblocked
                assert_unblocked(expected_window,record.get('window_scope'))
                item["input_attempted"] = result["input_attempted"] = True; checkpoint(result)
                if action == "press-key":item["execution"] = Controller.press_key(text)
                elif action == "hotkey":item["execution"] = Controller.hotkey(*text.split("+"))
                else:
                    item["execution"] = Controller.execute_resolved("set-value" if action == "set-text" else action,
                                                                     **target_arguments(control), text=text, **options)
                item["completed"] = True
                checkpoint(result)
            for check in step.get("checks", []):
                result["current_check"] = check["id"]; checkpoint(result)
                _control, evidence = observe(check, step.get("timeout_seconds", 10), assertion=True)
                checked={"id": check["id"], "kind": check["kind"],
                    "description": check.get("description") or check.get("target", {}).get("label", ""),
                    "expected": check.get("expected", True), **evidence}
                if after_check:
                    try:checked["observation"]=after_check(records[check["target"]["control_id"]],_control,main)
                    except Exception as error:checked["capture_error"]=str(error)
                item["checks"].append(checked);checkpoint(result)
            item["ok"] = True
            if after_action:
                try: item["after"] = after_action(records[step["target"]["control_id"]], main)
                except Exception as error: item["after_capture_error"] = str(error)
            result.pop("current_check", None); checkpoint(result)
        result.update(ok=True, message="流程执行及全部预期校验通过")
        return result
    except Exception as error:
        failed_check = result.get("current_check")
        if failed_check and result["steps"]:
            definition = next((c for s in flow["steps"] for c in s.get("checks", []) if c["id"] == failed_check), {})
            result["steps"][-1]["checks"].append({"id": failed_check, "kind": definition.get("kind"),
                "description": definition.get("description") or definition.get("target", {}).get("label", ""),
                "expected": definition.get("expected", True), "error": str(error),
                **getattr(error, "assertion_evidence", {}), "ok": False})
        result.update(error=str(error), code=getattr(error, "code", "action_failed"),
                      failed_step=result["steps"][-1]["id"] if result["steps"] else "precondition",
                      failed_check=result.get("current_check", ""))
        if getattr(error,'popup_state',None):result['popup_state']=error.popup_state
        if after_action and main is not None:
            try: result["failure_observation"] = after_action({}, main)
            except Exception as capture_error: result["failure_capture_error"] = str(capture_error)
        error.execution_result = result
        try: checkpoint(result)
        except Exception as checkpoint_error: error.add_note(str(checkpoint_error))
        raise
