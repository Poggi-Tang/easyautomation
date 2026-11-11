"""Immutable UI frames and their coordinate mapping, separate from action results."""

from __future__ import annotations

import time

from . import image_storage, knowledge


def _check_foreground(main_window):
    handle = int(getattr(main_window, "NativeWindowHandle", 0) or 0)
    if not handle:
        return
    import uiautomation
    from ..library_actions import owned_windows
    from .results import ActionError

    allowed = {handle, *(w["handle"] for w in owned_windows(
        {"handle": handle, "process_id": main_window.ProcessId}))}
    if int(uiautomation.GetForegroundWindow() or 0) not in allowed:
        raise ActionError("window_invalid", "前台已切换到其他软件，未保存其画面")


def describe(context_image: str, bounds: dict, control_rect: dict, main_bounds: dict,
             *, phase: str) -> dict:
    return {"phase": phase, "at": knowledge.utc_now(), "context_image": context_image,
            "context_size": {"width": bounds["width"], "height": bounds["height"]},
            "normalized_rect": knowledge.normalized_rect(control_rect, bounds),
            "main_rect": knowledge.normalized_rect(main_bounds, bounds)}


def union_bounds(main: dict, target: dict) -> dict:
    left, top = min(main["left"], target["left"]), min(main["top"], target["top"])
    right, bottom = max(main["right"], target["right"]), max(main["bottom"], target["bottom"])
    return dict(left=left, top=top, right=right, bottom=bottom, width=right-left, height=bottom-top)


def capture(directory, control, main_window, *, phase: str):
    """Capture the visible target and its context from the same frame before input."""
    from . import scanner
    from .results import ActionError
    from ..locator_images import ImageTarget, target_evidence

    main = scanner._window_rect(main_window)
    visual=isinstance(control,ImageTarget)
    rect = control.rectangle if visual else scanner._rect(control)
    if bool(getattr(control, "IsOffscreen", False)) or rect["width"] < 2 or rect["height"] < 2:
        raise ActionError("control_not_found", "控件已消失，未保存控件截图")
    bounds = union_bounds(main, rect)
    image = scanner._screenshot_window(bounds)
    if not visual and (scanner._rect(control) != rect or bool(getattr(control, "IsOffscreen", False))):
        raise ActionError("control_not_found", "控件在截图时发生变化，请重新定位")
    if visual:
        from PIL import Image
        with Image.open(control.reference) as reference:
            match=scanner.locate_template(reference,image,bounds)
        if not match["found"] or match["rect"]!=rect:
            raise ActionError("control_not_found","图片目标在截图时发生变化，未发送输入",reason="image_target_changed")
    _check_foreground(main_window)
    ref = image_storage.save_interaction_image(directory, image)
    crop = image.crop((rect["left"]-bounds["left"], rect["top"]-bounds["top"],
                       rect["right"]-bounds["left"], rect["bottom"]-bounds["top"]))
    return {**describe(ref, bounds, rect, main, phase=phase),**target_evidence(control)}, crop


def capture_after(directory, main_window) -> dict:
    """Wait briefly for compositor/menu animation without querying the old target."""
    from . import scanner

    started = stable_since = time.monotonic()
    previous, settled = None, False
    while True:
        _check_foreground(main_window)
        bounds = scanner._window_rect(main_window)
        if bounds["width"] < 2 or bounds["height"] < 2:
            raise ValueError("操作后主窗口已不可见，未生成结果截图")
        image = scanner._screenshot_window(bounds)
        signature = (tuple(bounds[k] for k in ("left", "top", "right", "bottom")), image.tobytes())
        now = time.monotonic()
        if signature != previous:
            stable_since = now
            previous = signature
        if now-started >= .2 and now-stable_since >= .12:
            settled = True
            break
        if now-started >= .8:
            break
        time.sleep(.04)
    _check_foreground(main_window)
    return {"phase": "after-action", "at": knowledge.utc_now(),
            "purpose": "execution-evidence",
            "settled": settled, "settle_seconds": round(time.monotonic()-started, 3),
            "context_image": image_storage.save_interaction_image(directory, image),
            "context_size": {"width": bounds["width"], "height": bounds["height"]},
            "main_rect": {"left": 0., "top": 0., "right": 1., "bottom": 1.}}


def reframe(rect: dict, main_rect: dict) -> dict:
    """Map main-window normalized geometry into an observation frame."""
    return {key: main_rect["left"] + rect[key]*(main_rect["right"]-main_rect["left"])
            if key in {"left", "right"} else main_rect["top"] + rect[key]*(main_rect["bottom"]-main_rect["top"])
            for key in ("left", "top", "right", "bottom")}
