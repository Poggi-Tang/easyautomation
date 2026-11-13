"""Region navigation and guards, executed only in bounded library workers."""
from __future__ import annotations
import time

from .mcp import native_windows
from .mcp.results import ActionError


def target_window(expected, scope):
    from .popup_guard import assert_unblocked
    assert_unblocked(expected,scope)
    if scope.get("kind", "main") == "main":return expected
    from .library_actions import owned_windows
    windows=[w for w in owned_windows(expected) if w["class_name"]==scope.get("class_name")
             and (not scope.get("title") or w["title"]==scope["title"])]
    if not windows:raise ActionError("control_not_found","区域所在对话框尚未显示",retryable=True)
    if len(windows)!=1:raise ActionError("ambiguous_match","多个对话框符合区域的窗口信息")
    return windows[0]


def condition_matches(condition, expected, main):
    from .library_actions import resolve_record
    if not condition:return False
    try:
        control=resolve_record(condition, expected, main)
    except ActionError as error:
        if error.code=="control_not_found" and not error.reason:return False
        raise
    if condition.get("kind", "exists")=="selected":
        pattern=control.GetSelectionItemPattern()
        if pattern is None:raise ValueError("状态条件控件不提供选中状态，请重新选择")
        return bool(pattern.IsSelected)
    return True


def open_region(library, app_id, region, expected, *, progress, checkpoint=lambda _r:None):
    from .library_actions import bind_window, resolve_record
    from .ctrl import Controller
    from . import workflows
    main=bind_window(expected)
    state=library.record_state(app_id,region);steps=library.entry_steps(app_id,region);condition=state.get("condition")
    execution=None
    if steps and not condition_matches(condition,expected,main):
        controls=library.read(app_id)["controls"]
        source=next((r for r in controls if r["id"]==region.get("source_control_id")),None)
        if source:steps=[*steps,workflows.click_step(source,action="inspect")]
        execution=workflows.execute(dict(keyword="打开区域",parameters={},steps=steps),controls,
                                    expected,progress=progress,checkpoint=checkpoint,skip_open_entries=True)
    if condition:
        deadline=time.monotonic()+3
        while not condition_matches(condition,expected,main):
            progress("等待区域状态："+condition.get("label",""))
            if time.monotonic()>=deadline:
                raise ValueError("前置步骤结束后，区域状态条件仍未满足；已发送的操作不会重试")
            time.sleep(.05)
    scope=region.get("window_scope") or library.scoped(region.get("page_id",library.page_id)).page(app_id).get("window_scope",{"kind":"main"})
    target=target_window(expected,scope)
    if region.get("location"):
        resolve_record({**region,"window_scope":scope},expected,main)
    Controller.activate_window_handle(target["handle"])
    return dict(ok=True,target_window=target,execution=execution,message="区域已打开")


def visible_region(library, app_id, region, expected):
    """Basic learning observes the current state and never replays entry steps."""
    from .library_actions import bind_window, resolve_record
    main = bind_window(expected)
    condition = library.record_state(app_id, region).get("condition")
    if condition and not condition_matches(condition, expected, main):
        raise ValueError("请先手动打开该区域的状态，再进行基本学习；基本学习不会执行前置步骤")
    scope = region.get("window_scope") or library.scoped(region.get("page_id",library.page_id)).page(app_id).get("window_scope", {"kind":"main"})
    target = target_window(expected, scope)
    if region.get("location"):resolve_record({**region,"window_scope":scope}, expected, main)
    return target
