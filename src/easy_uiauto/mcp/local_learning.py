"""Local annotations followed by the recorder's point-to-UIA binding.

UIA ancestors describe locations; they are not additional detected controls.
"""
from __future__ import annotations

import math
from pathlib import Path
import shutil
from tempfile import TemporaryDirectory

from . import knowledge, manual_layout


# These UIA nodes represent one widget. A canvas/list/custom node can contain
# several independent visual targets, so sharing those nodes is not sufficient.
_ATOMIC_CONTROLS = {
    "ButtonControl", "SplitButtonControl", "CheckBoxControl", "RadioButtonControl",
    "MenuItemControl", "HyperlinkControl", "EditControl", "ComboBoxControl",
    "SliderControl", "SpinnerControl", "TextControl", "TabItemControl",
}


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
            targets.append({**element, "id": len(targets)+1, "relative_rect": relative,
                "region_id": area["id"], "semantic_name": text or element.get("semantic_name") or "图形控件",
                "observed_value": text, "role": "label" if is_text(element) else "control",
                "source": "local-yolo-rapidocr", "evidence": [element.get("source", "local detection")],
                "analysis_result": reports[-1]["result"], "visual_components": []})
    return targets, reports


def resolve(window, bounds, targets, limit, progress):
    """Bind only annotated positions, using exactly the recording point method."""
    from .. import utils
    from . import scanner
    from .diagram_geometry import is_text

    controls, semantics, failures, paths, seen = [], {}, [], {}, {}
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
            kind = getattr(control, "ControlTypeName", "")
            if not belongs:
                reason = "outside_selected_window"
            elif bool(getattr(control, "IsOffscreen", False)):
                reason = "not_visible"
            elif kind in {"WindowControl", "GroupControl", "PaneControl", "ToolBarControl", "MenuBarControl", "TitleBarControl", "MenuControl"}:
                reason = "layout_container"
            elif not scanner._candidate_fits_target(control, absolute, window, details):
                reason = "geometry_rejected"
        if reason:
            failures.append(dict(error=reason, reason=reason, diagnostics=details, observation=target))
            continue
        key = scanner._control_key(control)
        if key in seen:
            index = seen[key]
            primary = semantics[index]
            components = primary["visual_components"]
            # Same-widget identity also covers a tiny icon nested in an outer
            # proposal (low IoU). Do not apply that shortcut to aggregate nodes.
            if (kind in _ATOMIC_CONTROLS
                    or any(scanner._same_visual_target(component, target) for component in components)):
                components.append({**target, "native_evidence": True})
                if is_text(target) and not primary.get("observed_value"):
                    primary["observed_value"] = target.get("observed_value", "")
                if is_text(primary) and not is_text(target):
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
        semantics[index] = {**target, **semantic, "entity_type": target.get("entity_type", "ui-element"),
                            "observed_value": target.get("observed_value", ""),
                            "visual_components": [{**target, "native_evidence": True}],
                            "analysis_result": target["analysis_result"]}
    return controls, semantics, failures, paths


def expanded_regions(directory, controls, context, bounds, frame, app_id, page_id):
    """Retain detected menu ancestors as regions, never extra control records."""
    from ..utils import get_control_xpath
    from . import scanner, control_observation
    from .protocol import location_from_xpath

    seen = set()
    for control, _depth in controls:
        current = control.GetParentControl()
        for _ in range(64):
            if current is None: break
            if getattr(current,"ControlTypeName","") == "MenuControl":
                xpath=get_control_xpath(current)
                identifier=knowledge.stable_id("menu-region",app_id,page_id,context["path"],xpath)
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
                    "source":"local-visual-container","coordinate_space":"observation",
                    "parent_control_id":context["path"][-1],"navigation_path":context["path"],
                    "location":location_from_xpath(xpath),"window_scope":context.get("scope",{"kind":"main"}),
                    "geometry_role":"visual-hint-only","normalized_rect":knowledge.normalized_rect(rectangle,capture_bounds),
                    "visual_observation":control_observation.describe(reference,capture_bounds,rectangle,bounds,phase="located"),
                    "image":str(destination.relative_to(directory)).replace("\\","/")})
                break
            current=current.GetParentControl()


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
        knowledge.write_markdown(target,preserve(metadata),body)
        path.unlink();archived+=1
    return archived
