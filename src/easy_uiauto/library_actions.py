"""Replay explicit knowledge navigation paths through the shared UIA controller."""

from __future__ import annotations

import os
import time

from .ctrl import Controller
from .mcp import knowledge, native_windows, scanner
from .mcp.protocol import normalize_location, normalize_match_mode
from .mcp.results import ActionError
from .utils import CONTROL_TYPE_IDS, _selector_index, observe_control
from .locator_images import ImageTarget, is_image_location, record_location, reference_images, target_arguments, target_evidence


def validate_plan(controls: list[dict], control_id: str, path=None) -> list[dict]:
    records = {r["id"]: r for r in controls}
    if control_id not in records:
        raise ValueError("目标控件已不存在，请重新学习")
    target = records[control_id]
    path = target.get("navigation_path", []) if path is None else path
    if (not isinstance(path, list) or len(path) > 20
            or any(not isinstance(i, str) or not i for i in path)):
        raise ValueError("前置步骤必须为最多 20 个控件 ID 的有序列表")
    identifiers = [*path, control_id]
    if len(set(identifiers)) != len(identifiers):
        raise ValueError("前置步骤不能重复或包含目标控件自身")
    plan = []
    for identifier in identifiers:
        record = records.get(identifier)
        if record is None:
            raise ValueError(f"前置控件已删除或尚未学习：{identifier}")
        if knowledge.runtime_quarantined(record):
            raise ActionError("operation_blocked", f"控件已因运行失败隔离：{identifier}")
        location = record_location(record)
        if not location["Xpath"] and not is_image_location(location):
            raise ValueError(f"控件缺少 XPath：{identifier}")
        if is_image_location(location) and location["PARAMETERS"].get("match_mode","identity")!="identity":
            raise ActionError("invalid_argument","图片定位不能使用同类序号的位置模式")
        default_mode = normalize_match_mode(location["PARAMETERS"].get("match_mode", "identity"))
        for step in [location, *location["Xpath"]]:
            _selector_index(step.get("foundIndex"))
            normalize_match_mode(step.get("match_mode", default_mode))
            if step.get("ControlType") and step["ControlType"] not in {*CONTROL_TYPE_IDS, "DesktopControl"}:
                raise ActionError("invalid_argument", "Unknown ControlType: "+str(step["ControlType"]))
            depth = step.get("searchDepth")
            if depth is not None:
                try:
                    if isinstance(depth, bool) or not isinstance(depth, (str, int)) or int(depth) < 0:
                        raise ValueError
                except ValueError as error:
                    raise ActionError("invalid_argument", "searchDepth must be a nonnegative integer") from error
        scope = record.get("window_scope", {"kind": "main"})
        if not isinstance(scope, dict) or scope.get("kind") not in {"main", "owned"}:
            raise ValueError(f"控件窗口上下文无效：{identifier}")
        if scope["kind"] == "owned" and not scope.get("class_name"):
            raise ValueError(f"弹出窗口缺少类名：{identifier}")
        plan.append({**record, "location": location, "window_scope": scope})
    return plan


def bind_window(expected: dict):
    """Validate runtime identity before binding UIA or changing focus."""
    import uiautomation

    candidates = native_windows.matching_windows("", handle=int(expected["handle"]))
    if (len(candidates) != 1 or not candidates[0]["visible"]
            or candidates[0]["process_id"] != expected["process_id"]
            or candidates[0]["class_name"] != expected["class_name"]):
        raise ActionError("window_invalid", "所选软件窗口已经失效，请重新选择")
    window = uiautomation.ControlFromHandle(expected["handle"])
    if window is None:
        raise ActionError("window_invalid", "无法绑定所选软件的 UI Automation")
    return window


def owned_windows(expected: dict, *, entries=None) -> list[dict]:
    windows = native_windows.enumerate_windows() if entries is None else entries
    by_handle = {w["handle"]: w for w in windows}

    def related(window):
        if window["process_id"] != expected["process_id"]:
            return False
        owner, seen = window.get("owner_handle", 0), set()
        if not owner and window.get("class_name") == "#32768":
            owner = native_windows.menu_owner(window["handle"])
        while owner and owner not in seen:
            if owner == expected["handle"]:
                return True
            seen.add(owner)
            owner = by_handle.get(owner, {}).get("owner_handle", 0)
        return False

    return [w for w in windows if w.get("visible", True) and w["handle"] != expected["handle"] and related(w)]


def activate_application(expected):
    """Keep an existing modal dialog active when its main window is disabled."""
    windows=owned_windows(expected)
    main=native_windows.matching_windows("",handle=expected["handle"])
    if len(main)!=1 or any(main[0][key]!=expected[key] for key in ("process_id","class_name")):
        raise ActionError("window_invalid","所选软件窗口已经失效")
    foreground=next((w for w in windows if w.get("foreground") and w.get("enabled",True)),None)
    target=foreground or (next((w for w in windows if w.get("enabled",True)),None)
                          if not main[0].get("enabled",True) else main[0])
    if target is None:raise ActionError("window_invalid","软件没有可激活的主窗口或所属对话框")
    return Controller.activate_window_handle(target["handle"])


def diagnose_uia_error(error, *, record=None, expected=None, root=None, uia_operation=False):
    """Explain permission failures without retrying or changing a saved locator."""
    current=error;seen=set()
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        hresult=getattr(current,"hresult",None)
        if (uia_operation and getattr(current,"winerror",None)==5 or
                isinstance(hresult,int) and hresult & 0xffffffff==0x80070005):
            title=(expected or {}).get("title") or "目标软件"
            return ActionError("operation_blocked",f"Windows 拒绝读取“{title}”的 UI Automation 控件。"
                "请让知识库管理器与目标软件使用相同的账户和权限；目标软件若以管理员身份运行，"
                "请以管理员身份重新启动管理器。已有区域和 XPath 已保留。",reason="uia_access_denied")
        current=current.__cause__
    if not isinstance(error,ActionError) or error.code!="control_not_found" or root is None:return error
    try:
        normalized=normalize_location((record or {}).get("location",{}))
        path=normalized["Xpath"]
        recorded_class=path[0].get("ClassName") if path else ""
        actual_class=getattr(root,"ClassName","")
    except Exception:return error
    if not recorded_class or not actual_class or actual_class==recorded_class:return error
    name=(record or {}).get("name") or normalized.get("Name") or "目标控件"
    message=f"“{name}”定位失败：记录的 UIA 根类为 {recorded_class}，本次读取为 {actual_class}。"
    if (expected or {}).get("process_id"):
        try:
            current_elevated=native_windows.process_elevated(os.getpid())
            target_elevated=native_windows.process_elevated(expected["process_id"])
        except OSError as diagnostic:
            if getattr(diagnostic,"winerror",None)==5:
                message+="当前进程也无权读取目标进程的权限信息，请优先检查账户和权限是否一致。"
                message+="若目标软件以管理员身份运行，请以管理员身份重新启动知识库管理器。"
        except Exception:
            pass  # A diagnostic failure must not replace the original lookup error.
        else:
            if target_elevated and not current_elevated:
                return ActionError("operation_blocked",message+"目标软件以管理员身份运行，管理器为普通权限。"
                    "请以管理员身份重新启动管理器；现有区域和 XPath 已保留。",reason="uia_access_denied")
    message+="定位检查未发送输入，也未改写 XPath；请检查软件窗口、权限或重新拾取区域。"
    return ActionError(error.code,message,retryable=error.retryable,reason=error.reason)


def resolve_record(record: dict, expected: dict, main_window, *, require_visible=True, directory=None):
    import uiautomation
    from .popup_guard import assert_unblocked

    scope = record.get("window_scope", {"kind": "main"})
    assert_unblocked(expected,scope)
    root = main_window
    location = record_location(record)
    visual_only=is_image_location(location)
    if scope.get("kind") == "owned":
        candidates = [w for w in owned_windows(expected)
                      if w["class_name"] == scope["class_name"]
                      and (not scope.get("title") or w["title"] == scope["title"])]
        if not candidates:
            raise ActionError("control_not_found", "前置步骤尚未打开目标弹出窗口", retryable=True)
        if len(candidates) != 1:
            raise ActionError("ambiguous_match", "多个弹出窗口符合控件的上下文")
        root = uiautomation.ControlFromHandle(candidates[0]["handle"])
        if root is None:
            raise ActionError("control_not_found", "目标弹出窗口暂不可读取", retryable=True)
        # UIA can include the main window above an owned dialog in the saved
        # XPath. Rebase only at an exact, unique ancestor matching this root.
        normalized = normalize_location(location)
        path = normalized["Xpath"]
        matches = [i for i, step in enumerate(path) if step.get("ControlType") in {"WindowControl", "MenuControl"}
                   and scanner._matches_step(root, step)]
        if len(matches) != 1 and not visual_only:
            # The native window can close after enumeration; only a fresh
            # native absence makes this a transient miss. A live mismatch
            # remains an invalid locator and must not be hidden by fallback.
            if not any(w["handle"]==candidates[0]["handle"] for w in owned_windows(expected)):
                raise ActionError("control_not_found", "目标弹出窗口已关闭", retryable=True)
            raise ActionError("invalid_argument", "弹出窗口的 XPath 与已绑定上下文不一致")
        if not visual_only:
            offset = int(path[matches[0]].get("searchDepth", matches[0]+1))-1
            normalized["Xpath"] = [{**step, "searchDepth": int(step.get("searchDepth", i+matches[0]+1))-offset}
                                   for i, step in enumerate(path[matches[0]:])]
            location = normalized
    if visual_only:
        if directory is None:raise ActionError("control_not_found","图片定位需要当前软件的知识库目录")
        _selector_index(location.get("foundIndex"))
        if any(location.get(key) for key in ("Name","ClassName","ControlType","AutomationId","foundIndex")):
            raise ActionError("invalid_argument","图片定位不能混入未解析的 UIA 属性，请使用 XPath 定位或移除这些属性")
        if location["PARAMETERS"].get("match_mode","identity")!="identity":
            raise ActionError("invalid_argument","图片定位不能使用位置模式")
        return resolve_image(record,directory,expected,root,ActionError("control_not_found","此目标为图片定位"))
    try:
        control = observe_control(location, window=root)
    except ActionError as error:
        diagnostic=diagnose_uia_error(error,record={**record,"location":location},expected=expected,root=root,uia_operation=True)
        if diagnostic is not error:raise diagnostic from error
        if directory is None or error.code!="control_not_found" or error.reason:raise
        normalized=normalize_location(location)
        if (normalized["PARAMETERS"].get("match_mode")=="position" or
                any(s.get("match_mode")=="position" for s in normalized["Xpath"])):raise
        return resolve_image(record,directory,expected,root,error)
    if type(control) is uiautomation.Control:
        # Generic XPath lookup returns Control; retain its exact UIA element
        # while exposing type-specific patterns such as Edit.GetValuePattern.
        control = uiautomation.Control.CreateControlFromControl(control)
        if control is None:
            raise ActionError("control_not_found", "控件在绑定 UIA 类型时失效", retryable=True)
    if not require_visible:
        return control
    if bool(getattr(control, "IsOffscreen", False)):
        raise ActionError("control_not_found", "控件当前不可见", retryable=True)
    rectangle = scanner._rect(control)
    if rectangle["width"] < 2 or rectangle["height"] < 2:
        raise ActionError("control_not_found", "控件尚无可点击区域", retryable=True)
    return control


def resolve_image(record, directory, expected, root, location_error):
    """Use fresh window pixels only after an ordinary, valid UIA miss."""
    from PIL import Image, ImageStat
    templates=[]
    for path in reference_images(directory,record):
        try:
            with Image.open(path) as image:
                template=image.convert("RGB")
            if min(template.size)>=8 and ImageStat.Stat(template.convert("L")).stddev[0]>=2:
                templates.append((path,template))
        except (OSError,ValueError):continue
    if not templates:
        raise ActionError("control_not_found",f"{location_error.message}；没有可用的定位图片",
                          reason="reference_image_missing") from location_error
    handle=int(getattr(root,"NativeWindowHandle",0) or expected["handle"])
    menu=record.get("window_scope",{}).get("class_name")=="#32768"
    if menu:activate_application(expected)
    else:Controller.activate_window_handle(handle,control=root)
    import uiautomation
    allowed={handle,expected["handle"]} if menu else {handle}
    if int(uiautomation.GetForegroundWindow() or 0) not in allowed:
        raise ActionError("window_invalid","图片匹配前前台已切换，未发送输入")
    bounds=scanner._window_rect(root)
    screen=scanner._screenshot_window(bounds)
    if int(uiautomation.GetForegroundWindow() or 0) not in allowed:
        raise ActionError("window_invalid","截图期间前台已切换，未发送输入")
    best=None
    for path,template in templates:
        match=scanner.locate_template(template,screen,bounds)
        if best is None or match["score"]>best[1]["score"]:best=(path,match)
    path,match=best
    if not match["found"]:
        raise ActionError("control_not_found",f"UIA 未找到控件；图片匹配未通过：{match['detail']}",
                          reason="image_match_failed") from location_error
    return ImageTarget(match["rect"],path,match["score"],match.get("uniqueness",0.),
                       record.get("semantic_name") or record.get("name", ""))


def _wait_record(record, expected, window, progress, seconds=3, directory=None):
    deadline = time.monotonic()+seconds
    while True:
        progress(f"正在定位：{record.get('semantic_name') or record.get('name') or record['id']}")
        try:
            return resolve_record(record, expected, window, **({"directory":directory} if directory is not None else {}))
        except ActionError as error:
            if (error.code != "control_not_found" or error.reason
                    or time.monotonic() >= deadline):
                raise
            time.sleep(.05)


def execute_plan(plan: list[dict], expected: dict, *, progress=None, checkpoint=None,
                 before_click=None, after_click=None, directory=None) -> dict:
    """Locate before activation; resolve each later step after its predecessor.

    Only read-only lookup is polled. Input is sent once and is never retried.
    """
    import uiautomation

    progress = progress or (lambda _message: None)
    checkpoint = checkpoint or (lambda _result: None)
    result = {"ok": False, "steps": [], "input_attempted": False}
    try:
        window = bind_window(expected)
        # UIA misses may activate the validated window only to match a usable template.
        minimized = bool(uiautomation.IsIconic(expected["handle"]))
        options={"directory":directory} if directory is not None else {}
        first = (resolve_record(plan[0], expected, window, require_visible=False, **options) if minimized
                 else _wait_record(plan[0], expected, window, progress, seconds=0, directory=directory))
        if plan[0].get("window_scope",{}).get("kind")=="owned":
            from .region_actions import target_window
            target=target_window(expected,plan[0]["window_scope"])
            Controller.activate_window_handle(target["handle"])
        elif expected.get("process_id") and expected.get("class_name"):activate_application(expected)
        else:Controller.activate_window_handle(expected["handle"], control=window)
        for index, record in enumerate(plan):
            progress(f"执行步骤 {index+1}/{len(plan)}：{record.get('semantic_name') or record.get('name') or record['id']}")
            if index+1 < len(plan):
                try:
                    next_control = resolve_record(plan[index+1], expected, window)
                except ActionError as error:
                    if error.code != "control_not_found" or error.reason:
                        raise
                else:
                    # Clicking an already open menu entry would close it again.
                    if bool(getattr(next_control, "IsEnabled", True)):
                        result["steps"].append({"control_id": record["id"],
                            "name": record.get("semantic_name") or record.get("name"),
                            "action": "click", "input_attempted": False, "completed": False,
                            "skipped": True, "reason": "next-control-already-visible"})
                        checkpoint(result)
                        continue
            control = first if index == 0 and not minimized and not isinstance(first,ImageTarget) else _wait_record(record, expected, window, progress, directory=directory)
            foreground = int(uiautomation.GetForegroundWindow() or 0)
            allowed = {expected["handle"], *(w["handle"] for w in owned_windows(expected))}
            if foreground not in allowed:
                raise ActionError("window_invalid", "前台已切换到其他软件，停止后续步骤")
            step = {"control_id": record["id"], "name": record.get("semantic_name") or record.get("name"),
                    "action": "click", "input_attempted": False, "completed": False}
            result["steps"].append(step)
            step.update(target_evidence(control))
            if before_click is not None:
                step["before"] = before_click(record, control, window)
                checkpoint(result)
                if int(uiautomation.GetForegroundWindow() or 0) not in allowed:
                    raise ActionError("window_invalid", "截图期间前台切换到其他软件，未执行当前点击")
            from .popup_guard import assert_unblocked
            assert_unblocked(expected,record.get('window_scope'))
            step["input_attempted"] = True
            result["input_attempted"] = True
            checkpoint(result)
            step["execution"] = Controller.execute_resolved("click", **target_arguments(control))
            step["completed"] = True
            checkpoint(result)
            if after_click is not None:
                try:
                    step["after"] = after_click(record, window)
                except Exception as error:
                    # Result-image failure must never repeat an already sent input.
                    step["after_capture_error"] = str(error)
                checkpoint(result)
        result.update(ok=True, message="点击已执行：" + " → ".join(
            (s["name"] or s["control_id"]) + ("（已展开）" if s.get("skipped") else "") for s in result["steps"]))
        return result
    except Exception as error:
        result.update(error=str(error), code=getattr(error, "code", "action_failed"))
        if getattr(error,'popup_state',None):result['popup_state']=error.popup_state
        if after_click is not None and result["input_attempted"]:
            try:
                result["failure_observation"] = after_click(plan[-1], window)
            except Exception as capture_error:
                result["failure_capture_error"] = str(capture_error)
        error.execution_result = result
        try:
            checkpoint(result)
        except Exception as cleanup_error:
            error.add_note(f"Saving execution status also failed: {cleanup_error}")
        raise
