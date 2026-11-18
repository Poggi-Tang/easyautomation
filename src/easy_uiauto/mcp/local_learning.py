"""Bind local annotations, or enumerate visible items inside a bound popup."""
from __future__ import annotations

import math
import time
from pathlib import Path
import shutil
from tempfile import TemporaryDirectory

from . import knowledge, manual_layout


# These UIA nodes represent one widget. A canvas/list/custom container can contain
# several independent visual targets, so sharing those nodes is not sufficient.
_ATOMIC_CONTROLS = {
    "ButtonControl", "SplitButtonControl", "CheckBoxControl", "RadioButtonControl",
    "MenuItemControl", "HyperlinkControl", "EditControl", "ComboBoxControl",
    "SliderControl", "SpinnerControl", "TextControl", "TabItemControl",
    "ListItemControl", "TreeItemControl",
}


def _contains(outer, inner):
    return (outer["width"] > 1 and outer["height"] > 1
            and all(outer[k] <= inner[k] for k in ("left", "top"))
            and all(outer[k] >= inner[k] for k in ("right", "bottom")))


def _fits_atomic_target(control, absolute, window):
    from . import scanner
    return (getattr(control, "ControlTypeName", "") in _ATOMIC_CONTROLS - {"TextControl"}
            and _contains(scanner._rect(control), absolute)
            and _contains(scanner._rect(window), scanner._rect(control)))


def _passive_part(control):
    """Labels/icons can belong to an item; interactive descendants stay separate."""
    from . import scanner
    if getattr(control, "IsKeyboardFocusable", False):
        return False
    getter = getattr(control, "GetPattern", None)
    if getter:
        patterns = scanner.uiautomation.PatternId
        for name in ("InvokePattern", "TogglePattern", "SelectionItemPattern", "ValuePattern",
                     "RangeValuePattern", "ExpandCollapsePattern"):
            if getter(getattr(patterns, name)) is not None:
                return False
        legacy = getter(patterns.LegacyIAccessiblePattern)
        if legacy is not None and getattr(legacy, "DefaultAction", ""):
            return False
    return True


def _atomic_owner(control, xpath, absolute, window, passive_cache):
    """Use only the recorded ancestry, never merge neighboring objects by text."""
    from . import scanner
    if getattr(control, "ControlTypeName", "") not in {"TextControl", "ImageControl"}:
        return control, xpath
    current = control
    for offset in range(min(len(xpath), 8)):
        if current is None or scanner._control_key(current) == scanner._control_key(window):
            break
        step = xpath[-1-offset]
        if not scanner._matches_step(current, step) or getattr(current, "IsOffscreen", False):
            break
        kind = getattr(current, "ControlTypeName", "")
        if offset and kind in _ATOMIC_CONTROLS - {"TextControl"}:
            if _fits_atomic_target(current, absolute, window):
                return current, xpath[:-offset]
            break
        if kind not in {"TextControl", "ImageControl", "PaneControl", "GroupControl"}:
            break
        key = scanner._control_key(current)
        if key not in passive_cache:
            passive_cache[key] = _passive_part(current)
        if not passive_cache[key]:
            break
        current = current.GetParentControl()
    return control, xpath


def detect(image, bounds, segmentation, directory, progress, *, expanded=False):
    """Use the same quality settings as the image CLI, without a vision API."""
    from ..local_image_worker import analyze
    from .diagram_geometry import is_text

    areas = segmentation["regions"]
    if expanded:
        areas = [{"id": "unassigned", "rect": dict(left=0, top=0, right=image.width,
                   bottom=image.height, width=image.width, height=image.height)}]
    targets, reports = [], []
    for area in areas:
        rect = area["rect"]
        crop = image.crop(tuple(rect[k] for k in ("left", "top", "right", "bottom")))
        progress(f"本地标注区域：{area.get('name', area['id'])}（YOLO＋RapidOCR）")
        # Intermediate annotations are kept together, separately from reference crops.
        with TemporaryDirectory(prefix="easy-uiauto-region-") as temporary:
            source = Path(temporary) / "region.png"
            crop.save(source)
            result = analyze(source, directory / "analysis", progress=progress)
        reports.append({"region_id": area["id"], "offset": {"x": rect["left"], "y": rect["top"]},
                        "result": str(Path(result["coordinates_json"]).relative_to(directory)).replace("\\", "/"),
                        "annotated_image": str(Path(result["annotated_image"]).relative_to(directory)).replace("\\", "/"),
                        "elements": result["element_count"], "truncated": result.get("truncated", False)})
        for element in result["elements"]:
            r = element["rect"]
            values = [r[k] for k in ("left", "top", "right", "bottom")]
            if any(not isinstance(v, (int, float)) or not math.isfinite(v) for v in values):
                raise ValueError("本地标注返回了无效坐标")
            left, top = max(0, round(values[0])), max(0, round(values[1]))
            right, bottom = min(crop.width, round(values[2])), min(crop.height, round(values[3]))
            if right <= left or bottom <= top:
                continue
            relative = dict(left=left+rect["left"], top=top+rect["top"],
                            right=right+rect["left"], bottom=bottom+rect["top"],
                            width=right-left, height=bottom-top)
            text = str(element.get("observed_value") or "").strip()
            parts = []
            for part in element.get("visual_components", []):
                pr = part["rect"]
                parts.append({**part, "role": "label", "component_relation": "part-of",
                    "relative_rect": {**pr, "left": pr["left"]+rect["left"],
                                      "right": pr["right"]+rect["left"],
                                      "top": pr["top"]+rect["top"], "bottom": pr["bottom"]+rect["top"]},
                    "analysis_result": reports[-1]["result"]})
            targets.append({**element, "id": len(targets)+1, "relative_rect": relative,
                "region_id": area["id"], "semantic_name": text or element.get("semantic_name") or "图形控件",
                "observed_value": text, "role": "label" if is_text(element) else "control",
                "source": "local-yolo-rapidocr", "evidence": [element.get("source", "local detection")],
                "analysis_result": reports[-1]["result"], "visual_components": parts})
    return targets, reports


def resolve(window, bounds, targets, limit, progress):
    """Read native popup items, otherwise bind the annotated recorder positions."""
    from .. import utils
    from . import scanner
    from .diagram_geometry import is_text

    menu = resolve_menu(window, bounds, limit, progress)
    if menu is not None:
        return menu

    controls, semantics, failures, paths, seen = [], {}, [], {}, {}
    passive_cache = {}
    oversized_labels = []
    for number, target in enumerate(targets[:limit], start=1):
        progress(f"按标注坐标读取控件：{number}/{min(limit,len(targets))}")
        rect = target["relative_rect"]
        x = bounds["left"] + (rect["left"]+rect["right"])//2
        y = bounds["top"] + (rect["top"]+rect["bottom"])//2
        reason, details = "", {"sample_point": {"x": x, "y": y}}
        try:
            xpath, control = utils.get_control_info(x, y, raise_errors=True)
        except Exception as error:
            from .results import ActionError
            # Only target-local lookup failures may become visual observations.
            # Permission, window and interference failures must stop the job.
            local_failure = (isinstance(error, ValueError)
                             and str(error) == "Recorded control is no longer a child of its parent")
            local_failure |= isinstance(error, ActionError) and error.code == "control_not_found"
            local_failure |= getattr(error, "hresult", None) in {-2147220991, 0x80040201}  # UIA element unavailable
            if not local_failure:
                from ..library_actions import diagnose_uia_error
                diagnostic = diagnose_uia_error(error, uia_operation=True)
                if diagnostic is not error:
                    raise diagnostic from error
                raise
            details.update(error_type=type(error).__name__, detail=str(error))
            failures.append(dict(error=str(error), reason="point_binding_failed",
                                 diagnostics=details, observation=target))
            continue
        if control is None or not xpath:
            reason = "point_lookup_failed"
        else:
            # Require the same window ancestry, including process and native identity.
            current, belongs = control, False
            for _ in range(64):
                if current is None: break
                if scanner._control_key(current) == scanner._control_key(window):
                    belongs = True; break
                current = current.GetParentControl()
            absolute = {**rect, "left": bounds["left"]+rect["left"], "right": bounds["left"]+rect["right"],
                        "top": bounds["top"]+rect["top"], "bottom": bounds["top"]+rect["bottom"]}
            if belongs and not bool(getattr(control, "IsOffscreen", False)):
                original_xpath = xpath
                control, xpath = _atomic_owner(control, xpath, absolute, window, passive_cache)
                if xpath != original_xpath:
                    target = {**target, "component_binding": "passive-native-child",
                              "native_location": scanner.location_from_xpath(original_xpath)}
            kind = getattr(control, "ControlTypeName", "")
            if not belongs:
                reason = "outside_selected_window"
            elif bool(getattr(control, "IsOffscreen", False)):
                reason = "not_visible"
            elif kind in {"WindowControl", "GroupControl", "PaneControl", "ToolBarControl", "MenuBarControl", "TitleBarControl", "MenuControl"}:
                reason = "layout_container"
            elif not (scanner._candidate_fits_target(control, absolute, window, details)
                      or _fits_atomic_target(control, absolute, window)):
                reason = "geometry_rejected"
        if reason:
            failure = dict(error=reason, reason=reason, diagnostics=details, observation=target)
            failures.append(failure)
            # OCR can include padding around a small icon (e.g. a magnifier read
            # as "Q"). Defer it until another proposal confirms the same widget.
            if (reason == "geometry_rejected" and is_text(target)
                    and kind in _ATOMIC_CONTROLS - {"TextControl"}
                    and details.get("reason") == "insufficient_overlap"
                    and details.get("area_ratio", 0) >= .25
                    and _contains(absolute, scanner._rect(control))
                    and _contains(scanner._rect(window), absolute)):
                oversized_labels.append((scanner._control_key(control), failure))
            continue
        target = {**target, "component_relation": "label-of" if is_text(target) else "part-of"}
        key = scanner._control_key(control)
        if key in seen:
            index = seen[key]
            primary = semantics[index]
            components = primary["visual_components"]
            # Same-widget identity also covers a tiny icon nested in an outer
            # proposal (low IoU). Do not apply that shortcut to aggregate nodes.
            if (kind in _ATOMIC_CONTROLS
                    or any(scanner._same_visual_target(component, target) for component in components)):
                was_text_only = all(is_text(c) for c in components)
                components.append({**target, "native_evidence": True})
                components.extend(target.get("visual_components", []))
                if is_text(target) and not primary.get("observed_value"):
                    primary["observed_value"] = target.get("observed_value", "")
                if not is_text(target) and was_text_only and kind not in _ATOMIC_CONTROLS - {"TextControl"}:
                    primary["entity_type"] = target.get("entity_type", "ui-element")
            else:
                failures.append(dict(error="distinct visual entity shares a UIA control",
                                     reason="shared_native_control", observation=target))
            continue
        index = len(controls)+1
        seen[key] = index
        controls.append((control, max(0, len(xpath)-1)))
        paths[id(control)] = xpath
        candidate = dict(uia_name=getattr(control,"Name", ""), automation_id=getattr(control,"AutomationId", ""),
                         control_type=getattr(control,"ControlTypeName", ""), help_text=getattr(control,"HelpText", ""),
                         supported_actions=scanner._control_actions(getattr(control,"ControlTypeName", ""),
                                                                  bool(getattr(control,"IsEnabled", True))))
        semantic = manual_layout.native_semantics(candidate)
        semantics[index] = {**target, **semantic,
                            "entity_type": ("ui-element" if kind in _ATOMIC_CONTROLS - {"TextControl"}
                                            else target.get("entity_type", "ui-element")),
                            "observed_value": (getattr(control,"Name", "") if kind in {"ListItemControl","TreeItemControl"}
                                               else target.get("observed_value", "")),
                            "visual_components": [{**target, "native_evidence": True},
                                                  *target.get("visual_components", [])],
                            "analysis_result": target["analysis_result"]}
    for key, failure in oversized_labels:
        if key not in seen:
            continue
        target = failure["observation"]
        primary = semantics[seen[key]]
        primary["visual_components"].append({**target, "native_evidence": True,
            "component_relation": "label-of", "component_binding": "same-native-widget"})
        failures.remove(failure)
    return controls, semantics, failures, paths


def _menu_window(window):
    kind = getattr(window, "ControlTypeName", "")
    class_name = getattr(window, "ClassName", "") or ""
    return (kind == "MenuControl" or class_name in {"#32768", "KMenu", "QMenu", "QComboBoxPrivateContainer"}
            or class_name.startswith("Qt") and "Popup" in class_name)


def resolve_menu(window, bounds, limit, progress):
    """Read the live menu subtree when point hit-testing only returns its popup.

    The caller has already bound and checked this native window. Do not search
    the desktop or the application's full tree, or import historical controls.
    """
    from . import scanner

    if not _menu_window(window):
        return None

    def step(control, depth, ordinal=None):
        value = dict(ControlType=control.ControlTypeName, searchDepth=depth, match_mode="identity")
        for key in ("Name", "ClassName", "AutomationId"):
            if getattr(control, key, ""):
                value[key] = getattr(control, key)
        if ordinal is not None:
            value["foundIndex"] = ordinal
        return value

    controls, semantics, failures, paths = [], {}, [], {}
    pending = [(window, [step(window, 1)])]
    seen = set()
    deadline = time.monotonic() + 3
    process_id = getattr(window, "ProcessId", 0)
    while pending:
        if len(seen) >= 2000 or time.monotonic() >= deadline or len(controls) >= limit:
            failures.append(dict(reason="menu_scan_truncated", error="当前菜单读取达到上限，未确认完整覆盖"))
            break
        control, path = pending.pop()
        key = scanner._control_key(control)
        if key in seen:
            continue
        seen.add(key)
        if bool(getattr(control, "IsOffscreen", False)):
            continue
        if process_id and getattr(control, "ProcessId", process_id) != process_id:
            continue
        kind = getattr(control, "ControlTypeName", "")
        if kind in {"MenuItemControl", "ListItemControl"}:
            rect = scanner._rect(control)
            if rect["width"] <= 1 or rect["height"] <= 1 or scanner._clip_rect(rect, bounds) != rect:
                continue
            progress(f"读取当前菜单项：{getattr(control, 'Name', '')}")
            relative = {**rect, "left": rect["left"]-bounds["left"], "right": rect["right"]-bounds["left"],
                        "top": rect["top"]-bounds["top"], "bottom": rect["bottom"]-bounds["top"]}
            native = manual_layout.native_semantics(dict(
                uia_name=getattr(control, "Name", ""), automation_id=getattr(control, "AutomationId", ""),
                control_type=kind, help_text=getattr(control, "HelpText", ""),
                supported_actions=scanner._control_actions(kind, bool(getattr(control, "IsEnabled", True)))))
            controls.append((control, len(path)-1))
            semantics[len(controls)] = {**native, "relative_rect": relative, "region_id": "unassigned",
                "observed_value": getattr(control, "Name", ""), "visual_components": [],
                "detection_method": "popup-menu-uia"}
            paths[id(control)] = path
            # A submenu is observed separately after opening it; hidden children
            # of its entry must not be presented as current executable controls.
            continue
        if kind not in {"MenuControl", "WindowControl", "PaneControl", "GroupControl", "CustomControl", "ListControl"}:
            continue
        children = control.GetChildren()
        if len(path) >= 8 and children:
            failures.append(dict(reason="menu_scan_truncated", error="当前菜单层级超限，未确认完整覆盖"))
            continue
        counts, descendants = {}, []
        for child in children:
            child_kind = child.ControlTypeName
            counts[child_kind] = counts.get(child_kind, 0) + 1
            descendants.append((child, [*path, step(child, len(path)+1, counts[child_kind])]))
        pending.extend(reversed(descendants))
    return (controls, semantics, failures, paths) if controls or failures else None


def expanded_regions(directory, controls, context, bounds, frame, app_id, page_id, *, paths=None, scopes=None):
    """Retain detected menu ancestors as regions, never extra control records."""
    from ..utils import get_control_xpath
    from . import scanner, control_observation
    from .protocol import location_from_xpath

    seen, assignments = set(), {}
    for control, _depth in controls:
        current = control.GetParentControl()
        for _ in range(64):
            if current is None: break
            if _menu_window(current):
                path=(paths or {}).get(id(control), [])
                matches=[i for i, item in enumerate(path) if item.get("ControlType")==current.ControlTypeName
                         and scanner._matches_step(current,item)]
                xpath=path[:matches[-1]+1] if matches else get_control_xpath(current)
                identifier=knowledge.stable_id("menu-region",app_id,page_id,context["path"],xpath)
                assignments[id(control)] = identifier
                if identifier in seen:break
                seen.add(identifier)
                rectangle=scanner._rect(current)
                capture_bounds,image,reference=frame(rectangle)
                crop=scanner._crop_control(image,rectangle,capture_bounds)
                image_path=knowledge.stable_id("region-image",identifier)+".png"
                destination=directory/"images/regions"/image_path;destination.parent.mkdir(parents=True,exist_ok=True)
                crop.save(destination)
                old=next((r for _p,r in knowledge.iter_records(directory,"region") if r["id"]==identifier),{})
                knowledge.save_region(directory,{**old,"id":identifier,"app_id":app_id,"page_id":page_id,
                    "name":old.get("name") or (getattr(current,"Name","") or "展开菜单"),
                    "source":"local-visual-container","coordinate_space":"main-window",
                    "parent_control_id":context["path"][-1],"navigation_path":context["path"],
                    "location":location_from_xpath(xpath),
                    "window_scope":(scopes or {}).get(id(control),(None,context.get("scope",{"kind":"main"})))[1],
                    "geometry_role":"visual-hint-only","normalized_rect":knowledge.normalized_rect(rectangle,bounds),
                    "visual_observation":control_observation.describe(reference,capture_bounds,rectangle,bounds,phase="located"),
                    "image":str(destination.relative_to(directory)).replace("\\","/"),
                    **({"entry_steps":context["entry_steps"]} if "entry_steps" in context else {})})
                break
            current=current.GetParentControl()
    return assignments


def matching_popup_record(records, location, scope, navigation):
    """Reuse a uniquely identified popup item, including a former standalone page.

    QAction names/IDs alone are not an identity: a toolbar can expose the same
    action. Require the owned window and the entire path within that popup.
    """
    from .control_hierarchy import _path
    wanted = _path({"location": location})
    matches = []
    for record in records:
        old_scope = record.get("window_scope", {})
        path = _path(record)
        tail=path[-len(wanted):] if wanted and len(path)>=len(wanted) else ()
        # The old path may include main-window ancestors. Its popup-root ordinal
        # is relative to those ancestors, not part of the standalone popup path.
        same_path=bool(tail and tail[0][:4]==wanted[0][:4] and tail[1:]==wanted[1:])
        legacy_popup = bool(same_path and (tail[0][0] == "MenuControl" or tail[0][3] in {
            "KMenu", "QMenu", "QComboBoxPrivateContainer", scope.get("class_name")}))
        if old_scope != scope and not (old_scope.get("kind", "main") == "main" and legacy_popup):continue
        old_navigation = record.get("navigation_path", [])
        if old_navigation and old_navigation != navigation:
            continue
        if same_path:
            matches.append(record)
    if len(matches) > 1:
        raise ValueError("菜单项有多个相同弹窗定位的历史记录，未创建或覆盖控件")
    return matches[0] if matches else {}


def archive_layout_records(directory, previous, observed, region_ids, scan_id):
    """Reversible migration of unused records from the former full-tree GUI scan."""
    containers={"GroupControl","PaneControl","ToolBarControl","TitleBarControl","MenuBarControl"}
    archived=0
    for record in previous:
        if (record["id"] in observed or record.get("control_type") not in containers
                or record.get("scan_strategy") != "manual-regions" or record.get("actions")
                or record.get("semantic_source") != "uia-metadata" or record.get("region_id") not in region_ids
                or record.get("navigation_path") or record.get("entry_steps") or record.get("region_behavior")
                or record.get("hierarchy_parent","auto") != "auto"):
            continue
        identifier=record["id"]
        found=next(((p,r) for p,r in knowledge.iter_records(directory,"control") if r["id"]==identifier),None)
        if found is None:continue
        path,_record=found
        # Include workflows and hand-written guides, not only the control index.
        if any(identifier in p.read_text(encoding="utf-8") for p in directory.rglob("*.md")
               if p!=path and ".easy_uiauto" not in p.relative_to(directory).parts and p.name not in {"index.md","commands.md"}):
            continue
        archive_record(directory, path, scan_id)
        archived+=1
    return archived


def archive_record(directory, path, scan_id, *, reason=None):
    """Keep the original record and its images before removing it from the tree."""
    target=directory/".easy_uiauto/archive"/scan_id/"controls"/path.name
    target.parent.mkdir(parents=True,exist_ok=True)
    archive=target.parent.parent
    def preserve(value):
        if isinstance(value,dict):return {k:preserve(v) for k,v in value.items()}
        if isinstance(value,list):return [preserve(v) for v in value]
        if isinstance(value,str) and value.startswith("images/"):
            source=(directory/value).resolve()
            if source.is_relative_to(directory.resolve()) and source.is_file():
                destination=archive/value;destination.parent.mkdir(parents=True,exist_ok=True)
                shutil.copy2(source,destination)
                return destination.relative_to(directory).as_posix()
        return value
    metadata,body=knowledge.read_markdown(path)
    if reason:
        metadata["component_archive"] = reason
    knowledge.write_markdown(target,preserve(metadata),body)
    path.unlink()
