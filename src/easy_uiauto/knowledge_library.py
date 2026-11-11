"""Manage manual layouts in the existing Markdown UI knowledge vault."""

from __future__ import annotations

import json
import shutil
from pathlib import Path
from uuid import uuid4

from .mcp import dynamic_regions, image_storage, knowledge, manual_layout
from .mcp.protocol import normalize_location, with_match_mode

PAGE_ID = "manual-main"


class KnowledgeLibrary:
    def __init__(self, root: str | Path | None = None, *, page_id: str = PAGE_ID):
        if not page_id or knowledge.slugify(page_id) != page_id:
            raise ValueError("区域状态 ID 无效")
        self.page_id = page_id
        self.root = Path(root).expanduser().resolve() if root else knowledge.vault_root().resolve()

    def scoped(self, page_id: str):
        """A state uses its own existing page namespace; legacy IDs stay intact."""
        return KnowledgeLibrary(self.root, page_id=page_id)

    def state(self, app_id: str) -> dict:
        return self.page(app_id).get("region_state", {})

    def record_state(self, app_id: str, record: dict, *, snapshot=None) -> dict:
        page_id = record.get("page_id", self.page_id)
        regions = snapshot["regions"] if snapshot is not None else self.scoped(page_id).regions(app_id)
        region = record if record.get("kind") == "region" else next((r for r in regions
            if r["id"] in {record.get("region_id"), page_id+"."+str(record.get("region_id", ""))}), {})
        defaults = (next((p for p in snapshot["pages"] if p["id"]==page_id),{}).get("region_state",{})
                    if snapshot is not None else self.scoped(page_id).state(app_id))
        return {**defaults, "name": region.get("state_name", defaults.get("name", "")),
                "steps": region.get("entry_steps", defaults.get("steps", [])),
                "condition": region.get("state_condition", defaults.get("condition"))}

    def configure_region(self, app_id, region_id, *, steps, condition=None, state_name=""):
        from . import workflows
        records = self.read(app_id)
        region = next((r for r in records["regions"] if r["id"] == region_id), None)
        if region is None: raise ValueError("区域已不存在")
        if steps:
            report = workflows.readiness(dict(keyword="打开区域", parameters={}, steps=list(steps)), records["controls"])
            if not report["ready"]: raise ValueError("前置步骤尚不可执行："+"；".join(i["message"] for i in report["issues"]))
        if condition:
            condition = {**condition, "location": normalize_location(condition.get("location"))}
            if not condition["location"]["Xpath"] or condition.get("kind", "exists") not in {"exists", "selected"}:
                raise ValueError("状态条件无效")
        with image_storage._vault_lock(self.root):
            region.update(entry_steps=json.loads(json.dumps(list(steps))), state_condition=condition, state_name=state_name)
            knowledge.save_region(self.directory(app_id), region); self._refresh(self.directory(app_id))
        return region

    def configure_state(self, app_id: str, *, name: str, area_id: str = "", steps=(), condition=None):
        from . import workflows

        if not name.strip(): raise ValueError("状态名称不能为空")
        steps = json.loads(json.dumps(list(steps)))
        controls = self.read(app_id)["controls"]
        if steps:
            flow = dict(keyword="打开区域", parameters={}, steps=steps)
            report = workflows.readiness(flow, controls)
            if not report["ready"]:
                raise ValueError("前置步骤尚未配置完成：" + "；".join(i["message"] for i in report["issues"]))
        if condition:
            location = normalize_location(condition.get("location"))
            if not location["Xpath"] or condition.get("kind", "exists") not in {"exists", "selected"}:
                raise ValueError("状态条件需要有效定位，类型为存在或选中")
            condition = {**condition, "location": location}
        directory = self.directory(app_id)
        with image_storage._vault_lock(self.root):
            page = self.page(app_id)
            if not page: raise ValueError("请先选择区域")
            page["region_state"] = dict(id=self.page_id, name=name.strip(), area_id=area_id,
                                        steps=steps, condition=condition)
            knowledge.save_page(directory, page)
            self._refresh(directory)
        return page["region_state"]

    def save_variant(self, app_id: str, base_region: dict, result: dict, *, name: str, steps=(), condition=None):
        """Overlapping Tab views live in different namespaces, never a flat layout."""
        if not name.strip(): raise ValueError("请输入区域状态名称")
        if not base_region.get("id") or not any(r["id"] == base_region["id"] for r in self.read(app_id)["regions"]):
            raise ValueError("原区域已不存在")
        if not steps or not condition:
            raise ValueError("新状态需要前置步骤和明确的状态条件")
        # Validate before saving a new page or reference image.
        from . import workflows
        report = workflows.readiness(dict(keyword=name, steps=list(steps), parameters={}), self.read(app_id)["controls"])
        if not report["ready"]: raise ValueError("新状态的前置步骤尚未配置完成")
        normalized = normalize_location(condition.get("location"))
        if not normalized["Xpath"] or condition.get("kind", "exists") not in {"exists", "selected"}:
            raise ValueError("状态条件需要有效定位，类型为存在或选中")
        state_library = self.scoped("state-" + uuid4().hex[:12])
        identifier = state_library.save_picked_region(app_id, result, name=base_region.get("name", name))
        state_library.configure_state(app_id, name=name, area_id=self.area_id(app_id, base_region),
                                      steps=steps, condition=condition)
        return identifier

    def area_id(self, app_id: str, region: dict) -> str:
        return self.scoped(region.get("page_id", PAGE_ID)).state(app_id).get("area_id") or region["id"]

    def entry_steps(self, app_id: str, record: dict, *, snapshot=None) -> list[dict]:
        """Expose the actual route for single-control execution or explicit insertion."""
        from . import workflows
        if "entry_steps" in record:
            combined = record["entry_steps"]
            return self._fresh_steps(combined)
        state_steps = self.record_state(app_id, record,snapshot=snapshot).get("steps", [])
        entries = {r["id"]: r for r in (snapshot if snapshot is not None else self.read(app_id))["controls"]}
        identifiers=record.get("navigation_path", [])
        missing=[i for i in identifiers if i not in entries]
        if missing:raise ValueError("前置控件已删除或尚未学习："+", ".join(missing))
        paths = [{**workflows.click_step(entries[i]),"skip_when_next_visible":True} for i in identifiers]
        # Only merge a shared leading path; repeated later clicks remain meaningful.
        overlap = 0
        while overlap < min(len(state_steps), len(paths)) and state_steps[overlap].get("target") == paths[overlap].get("target"):
            overlap += 1
        combined = [*state_steps, *paths[overlap:]]
        return self._fresh_steps(combined)

    @staticmethod
    def _fresh_steps(steps):
        copied=json.loads(json.dumps(steps))
        for step in copied:
            step["id"]="entry-"+uuid4().hex[:10]
            for check in step.get("checks",[]):check["id"]="entry-check-"+uuid4().hex[:10]
        return copied

    def save_captured_control(self, app_id: str, result: dict, *, name="") -> dict:
        """Archive a picked entry before learning its destination; no live COM crosses here."""
        from PIL import Image
        selection, capture = result["selection"], result["capture"]
        location = normalize_location(selection["location"])
        if not location["Xpath"]: raise ValueError("所选控件没有 XPath")
        identifier = knowledge.stable_id("control", app_id, self.page_id, location, selection.get("window_scope", {}))
        directory = self.directory(app_id)
        with image_storage._vault_lock(self.root):
            n = selection["normalized_rect"]
            with Image.open(capture["image"]) as image:
                context_ref=image_storage.save_interaction_image(directory,image)
                context_size=dict(width=image.width,height=image.height)
                crop = image.crop(tuple(round(n[k] * (image.width if k in {"left", "right"} else image.height))
                                        for k in ("left", "top", "right", "bottom")))
                image_ref = image_storage.save_interaction_image(directory, crop)
            old = knowledge.find_control_record(directory, identifier, include_quarantine=True)
            record = {**(old[1] if old else {}), "id": identifier, "app_id": app_id, "page_id": self.page_id,
                      "name": selection["native_name"], "semantic_name": name or selection["native_name"] or selection["control_type"],
                      "control_type": selection["control_type"], "class_name": selection["class_name"],
                      "automation_id": selection["automation_id"], "location": location, "image": image_ref,
                      "normalized_rect": n, "region_id": "unassigned", "status": "observed",
                      "visual_observation":dict(phase="picked",purpose="learning-reference",context_image=context_ref,
                                                context_size=context_size,normalized_rect=n),
                      "window_scope": selection.get("window_scope", {"kind":"main"}),
                      "notes": "Captured prerequisite; business effect has not been verified."}
            knowledge.save_control(directory, record); self._refresh(directory)
        return record

    def directory(self, app_id: str) -> Path:
        if not app_id or knowledge.slugify(app_id) != app_id:
            raise ValueError("软件 ID 只能包含字母、数字、中文、点、连字符和下划线")
        directory = self.root / "applications" / app_id
        if (not directory.resolve().is_relative_to(self.root)
                or directory.is_symlink() or getattr(directory, "is_junction", lambda: False)()):
            raise ValueError("软件目录不能是符号链接或目录联接")
        for name in ("pages", "regions", "controls", "quarantine", "images", "workflows", "usage", "references", ".easy_uiauto"):
            child = directory / name
            if child.is_symlink() or getattr(child, "is_junction", lambda: False)():
                raise ValueError("知识库记录目录不能是符号链接或目录联接")
        return directory

    def apps(self) -> list[dict]:
        applications = self.root / "applications"
        if not applications.is_dir():
            return []
        result = []
        for directory in sorted(applications.iterdir()):
            path = directory / "app.md"
            if not directory.is_dir() or not path.is_file():
                continue
            try:
                self.directory(directory.name)
                metadata, _body = knowledge.read_markdown(path)
            except (OSError, ValueError):
                continue
            result.append({**metadata, "id": directory.name})
        return result

    def read(self, app_id: str) -> dict:
        directory = self.directory(app_id)
        if not (directory / "app.md").is_file():
            raise ValueError("软件尚未加入知识库")
        metadata, _body = knowledge.read_markdown(directory / "app.md")
        return {"schema_version": knowledge.INDEX_VERSION, "application": metadata,
                "workflows": self.workflows(app_id),
                "usage": self.usage_documents(app_id),
                "references": self.reference_documents(app_id),
                **{plural: [r for _p, r in knowledge.iter_records(directory, kind)]
                   for plural, kind in (("pages", "page"), ("regions", "region"),
                                        ("controls", "control"), ("interactions", "interaction"))}}

    def reference_documents(self, app_id):
        from . import reference_library
        return reference_library.documents(self.directory(app_id))

    def import_reference(self, app_id, source, **kwargs):
        from . import reference_library
        return reference_library.import_reference(self.directory(app_id), source, **kwargs)

    def save_reference(self, app_id, identifier, body, *, enabled=True):
        from . import reference_library
        reference_library.save(self.directory(app_id), identifier, body, enabled=enabled)

    def workflows(self, app_id: str) -> list[dict]:
        from . import workflows

        result = []
        directory = self.directory(app_id)
        controls = [r for _p, r in knowledge.iter_records(directory, "control")]
        for path in sorted((directory / "workflows").glob("*.md")):
            if image_storage._safe_path(directory, path) is None: raise ValueError("流程文件不能为链接")
            record, _body = knowledge.read_markdown(path)
            if record.get("kind") == "workflow":
                if record.get("last_run"):
                    record["last_run"] = {**record["last_run"], "outdated":
                        record["last_run"].get("definition_hash") != workflows.definition_hash(record)
                        or record["last_run"].get("control_hash") != workflows.control_hash(record, controls)}
                result.append(record)
        return result

    def usage_documents(self, app_id: str) -> list[dict]:
        """Handwritten instructions and existing page documentation, without migration."""
        directory = self.directory(app_id)
        result = []
        for path in [directory / "app.md", *sorted((directory / "pages").glob("*.md")),
                     *sorted((directory / "usage").glob("*.md"))]:
            if image_storage._safe_path(directory, path) is None:
                raise ValueError("使用知识文件不能为链接")
            metadata, body = knowledge.read_markdown(path)
            relative = path.relative_to(directory).as_posix()
            result.append({**metadata, "document_path": relative, "body": body,
                           "editable": relative.startswith("usage/"),
                           "name": metadata.get("name") or metadata.get("title") or path.stem})
        return result

    def save_usage(self, app_id: str, name: str, body: str, *, document_id="", workflow_ids=()) -> dict:
        if not name.strip(): raise ValueError("使用知识标题不能为空")
        directory = self.directory(app_id)
        identifier = document_id or "usage-" + uuid4().hex[:12]
        if knowledge.slugify(identifier) != identifier: raise ValueError("使用知识 ID 无效")
        path = directory / "usage" / (identifier + ".md")
        if image_storage._safe_path(directory, path) is None: raise ValueError("使用知识路径无效")
        with image_storage._vault_lock(self.root):
            available = {r["id"] for r in self.workflows(app_id)}
            if not isinstance(workflow_ids, (list, tuple)) or any(i not in available for i in workflow_ids):
                raise ValueError("引用的组合动作不属于当前软件或已删除")
            record = dict(id=identifier, kind="usage", app_id=app_id, name=name.strip(),
                          workflow_ids=list(dict.fromkeys(workflow_ids)), updated_at=knowledge.utc_now())
            knowledge.write_markdown(path, record, body)
        return {**record, "body": body, "document_path": path.relative_to(directory).as_posix(), "editable": True}

    def save_workflow(self, app_id: str, flow: dict) -> dict:
        from . import workflows

        workflows.validate(flow)
        directory = self.directory(app_id)
        identifier = flow.get("id") or knowledge.stable_id("workflow", flow["keyword"])
        if not isinstance(identifier, str) or knowledge.slugify(identifier) != identifier: raise ValueError("流程 ID 无效")
        path = directory / "workflows" / (identifier+".md")
        if image_storage._safe_path(directory, path) is None: raise ValueError("流程路径必须位于软件知识库中")
        with image_storage._vault_lock(self.root):
            if any(r["keyword"] == flow["keyword"] and r["id"] != identifier for r in self.workflows(app_id)):
                raise ValueError("本软件已有同名流程关键字")
            controls = self.read(app_id)["controls"]
            status = workflows.readiness(flow, controls)
            record = {**flow, "id": identifier, "kind": "workflow", "app_id": app_id,
                      "status": "ready" if status["ready"] else "draft", "readiness": status,
                      "definition_hash": workflows.definition_hash(flow), "updated_at": knowledge.utc_now()}
            if record.get("last_run"):
                record["last_run"] = {**record["last_run"], "outdated":
                    record["last_run"].get("definition_hash") != record["definition_hash"]
                    or record["last_run"].get("control_hash") != workflows.control_hash(record, controls)}
            body = f"# {record['keyword']}\n\n{record.get('name', '')}\n\n前置条件：{record.get('precondition', '')}\n\n"
            body += "\n".join(f"{n}. {s.get('description', s['action'])}\n   预期：{s.get('expected_description', '')}"
                              for n, s in enumerate(record["steps"], 1))
            knowledge.write_markdown(path, record, body)
            return record

    def import_workflow_case(self, app_id: str, markdown: str, keyword="") -> dict:
        from . import workflows

        flow = workflows.import_case(markdown, app_id, keyword)
        with image_storage._vault_lock(self.root):
            if any(r["id"] == flow["id"] or r["keyword"] == flow["keyword"] for r in self.workflows(app_id)):
                raise ValueError("已存在此流程关键字，请编辑已有流程，或输入一个新的关键字再转换")
            return self.save_workflow(app_id, workflows.suggest_bindings(flow, self.read(app_id)["controls"]))

    def save_workflow_execution(self, app_id: str, flow_id: str, execution: dict) -> dict:
        with image_storage._vault_lock(self.root):
            flow = next((r for r in self.workflows(app_id) if r["id"] == flow_id), None)
            if flow is None: raise ValueError("流程已不存在")
            return self.save_workflow(app_id, {**flow, "last_run": {**execution, "at": knowledge.utc_now()}})

    def bind(self, app_id: str, name: str, window: dict) -> Path:
        if not name.strip():
            raise ValueError("请输入软件名称")
        directory = self.directory(app_id)
        with image_storage._vault_lock(self.root):
            existing = directory / "app.md"
            previous = knowledge.read_markdown(existing)[0] if existing.is_file() else {}
            knowledge.initialize_app(app_id, name.strip(), self.root)
            record, body = knowledge.read_markdown(directory / "app.md")
            record = {**previous, **record}
            record.update(name=name.strip(), window_name=window["title"],
                          window_class=window["class_name"], layout_source="manual")
            knowledge.write_markdown(directory / "app.md", record, body)
            knowledge.rebuild_index(directory)
        return directory

    def asset(self, app_id: str, reference: str) -> Path | None:
        if not reference or not isinstance(reference, str):
            return None
        directory = self.directory(app_id).resolve()
        try:
            candidate = (directory / reference).resolve()
        except (OSError, ValueError):
            return None
        return candidate if candidate.is_relative_to(directory) and candidate.is_file() else None

    def page(self, app_id: str) -> dict:
        path = self.directory(app_id) / "pages" / f"{self.page_id}.md"
        return knowledge.read_markdown(path)[0] if path.is_file() else {}

    def save_capture(self, app_id: str, capture: dict) -> dict:
        directory = self.directory(app_id)
        with image_storage._vault_lock(self.root):
            previous = self.page(app_id)
            destination = directory / "images" / "pages" / f"{self.page_id}-layout.png"
            if image_storage._safe_path(directory, destination) is None:
                raise ValueError("截图路径必须位于当前软件目录内")
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(capture["image"], destination)
            size = {"width": capture["width"], "height": capture["height"]}
            changed = previous.get("observed_window_size") != size or previous.get("window_class") != capture["window_class"]
            record = {**previous, "id": self.page_id, "app_id": app_id, "name": "人工区域布局",
                      "window_name": capture["window_name"], "window_class": capture["window_class"],
                      "image": str(destination.relative_to(directory)).replace("\\", "/"),
                      "layout_image": str(destination.relative_to(directory)).replace("\\", "/"),
                      "observed_window_size": size, "region_source": "manual",
                      "requires_layout_review": bool(previous.get("requires_layout_review")
                                                     or (previous.get("observed_window_size") and changed)),
                      "layout_confirmed": bool(previous.get("layout_confirmed") and not changed)}
            if capture.get("window_scope"):
                record["window_scope"] = capture["window_scope"]
            knowledge.save_page(directory, record)
            knowledge.rebuild_index(directory)
        return record

    def regions(self, app_id: str) -> list[dict]:
        return [r for _p, r in knowledge.iter_records(self.directory(app_id), "region")
                if r.get("page_id") == self.page_id and r.get("source") == "manual"
                and not r.get("classification_only")]

    def layout_regions(self, app_id: str) -> list[dict]:
        return [{**r, "id": r["id"].removeprefix(self.page_id + ".")}
                for r in self.regions(app_id)]

    def sync_dynamic_regions(self, app_id: str) -> None:
        """Upgrade learned popup containers without performing any UI action."""
        directory = self.directory(app_id)
        with image_storage._vault_lock(self.root):
            if dynamic_regions.synchronize(directory, self.read(app_id)["controls"]):
                self._refresh(directory)

    def make_dynamic_region(self, app_id: str, control_id: str) -> str:
        directory = self.directory(app_id)
        with image_storage._vault_lock(self.root):
            found = knowledge.find_control_record(directory, control_id, include_quarantine=True)
            if not found or not found[1].get("navigation_path"):
                raise ValueError("请选择通过“学习展开控件”学到的容器，区域需要已保存的展开步骤")
            record = found[1]
            if record.get("control_type") not in {"MenuControl", "PaneControl", "GroupControl", "ToolBarControl",
                    "WindowControl", "CustomControl", "ListControl", "TreeControl", "TableControl", "DataGridControl"}:
                raise ValueError("请选择菜单、面板等容器；按钮和菜单项保留为操作控件")
            observation = record.get("visual_observation", {})
            if not self.asset(app_id, observation.get("context_image", "")) or not observation.get("normalized_rect"):
                raise ValueError("容器缺少完整的学习参考图，请先重新学习展开控件")
            record["region_behavior"] = "region"
            knowledge.save_control(directory, record)
            dynamic_regions.synchronize(directory, self.read(app_id)["controls"])
            self._refresh(directory)
        return knowledge.stable_id("dynamic-region", control_id)

    def save_dynamic_region(self, app_id: str, name: str, rectangle: dict, *,
                            parent_region_id: str = "", region_id: str = "", description: str = "") -> str:
        """Draw only inside a saved expanded-state frame, never on a result frame."""
        directory = self.directory(app_id)
        with image_storage._vault_lock(self.root):
            regions = {r["id"]: r for r in self.read(app_id)["regions"]}
            previous = regions.get(region_id, {})
            if region_id and not dynamic_regions.is_dynamic(previous):
                raise ValueError("动态区域不存在")
            parent = regions.get(parent_region_id or previous.get("parent_region_id"), {})
            reference = previous if previous.get("source") == "dynamic-container" else parent
            if not dynamic_regions.is_dynamic(reference):
                raise ValueError("请先选中一个已学习的动态区域")
            candidate = {"id": region_id or f"dynamic-region-{uuid4().hex[:12]}",
                         "name": name.strip(), "normalized_rect": rectangle}
            manual_layout.validate_regions([candidate])
            if previous.get("source") == "dynamic-container":
                # Native bounds belong to the source control; editing a region must
                # never reinterpret its locator or silently replace learned pixels.
                size = previous.get("visual_observation", {}).get("context_size", {})
                if any(abs(rectangle[k]-previous["normalized_rect"][k]) > 1/max(2, size.get(
                        "width" if k in {"left", "right"} else "height", 10000)) for k in rectangle):
                    raise ValueError("菜单容器边界来自 UIA；请在其中划分子区域，而不是修改容器边界")
                knowledge.save_region(directory, {**previous, "name": candidate["name"], "description": description})
                self._refresh(directory)
                return region_id
            bounds = reference.get("normalized_rect", {})
            if any(rectangle[k] < bounds[k]-1e-6 for k in ("left", "top")) or any(
                    rectangle[k] > bounds[k]+1e-6 for k in ("right", "bottom")):
                raise ValueError("子区域必须位于上级区域边界内")
            if previous.get("source") != "dynamic-container":
                siblings = [r for r in regions.values() if r.get("parent_region_id") == reference["id"]
                            and r["id"] != region_id]
                manual_layout.validate_regions([*siblings, candidate])
            for child in regions.values():
                if child.get("parent_region_id") == region_id:
                    n = child["normalized_rect"]
                    if any(n[k] < rectangle[k]-1e-6 for k in ("left", "top")) or any(
                            n[k] > rectangle[k]+1e-6 for k in ("right", "bottom")):
                        raise ValueError("新的边界不能排除已有子区域")
            observation = {**reference.get("visual_observation", {}), "normalized_rect": rectangle}
            source = self.asset(app_id, observation.get("context_image", ""))
            if source is None:
                raise ValueError("学习参考图缺失，请重新学习展开控件")
            from PIL import Image
            with Image.open(source) as image:
                pixels = tuple(round(rectangle[k]*(image.width if k in {"left", "right"} else image.height))
                               for k in ("left", "top", "right", "bottom"))
                if pixels[2]-pixels[0] < 2 or pixels[3]-pixels[1] < 2:
                    raise ValueError("区域宽高至少为 2 像素")
                image_ref = image_storage.save_interaction_image(directory, image.crop(pixels))
            record = {**previous, **candidate, "app_id": app_id, "page_id": reference["page_id"],
                      "description": description, "image": image_ref, "visual_observation": observation,
                      "geometry_role": "visual-hint-only", "coordinate_space": "observation"}
            if previous.get("source") != "dynamic-container":
                record.update(source="manual-dynamic", parent_region_id=reference["id"],
                              navigation_path=reference.get("navigation_path", []),
                              window_scope=reference.get("window_scope", {"kind": "main"}))
            knowledge.save_region(directory, record)
            self._refresh(directory)
            return record["id"]

    def save_region(self, app_id: str, name: str, rectangle: dict,
                    region_id: str = "", *, description: str = "", control_record: dict | None = None) -> str:
        directory = self.directory(app_id)
        page = self.page(app_id)
        if not page:
            raise ValueError("请先选择软件并获取截图")
        identifier = region_id.removeprefix(self.page_id + ".") if region_id else f"region-{uuid4().hex[:10]}"
        if knowledge.slugify(identifier) != identifier:
            raise ValueError("区域 ID 不合法")
        previous = next((r for r in self.layout_regions(app_id) if r["id"] == identifier), {})
        candidate = {**previous, "id": identifier, "name": name.strip(), "normalized_rect": rectangle,
                     "edited_by_user": True}
        if control_record is not None:
            candidate.update({k: control_record[k] for k in (
                "location", "automation_id", "native_name", "class_name", "control_type", "picked_image_size")})
            candidate.update(selection_source="uia-picker", geometry_role="uia-control-region")
        regions = [r for r in self.layout_regions(app_id) if r["id"] != identifier]
        manual_layout.validate_regions([*regions, candidate])
        with image_storage._vault_lock(self.root):
            image_ref = ""
            source = self.asset(app_id, page.get("layout_image", page.get("image", "")))
            if source:
                from PIL import Image

                with Image.open(source) as image:
                    bounds = (round(rectangle["left"] * image.width), round(rectangle["top"] * image.height),
                              round(rectangle["right"] * image.width), round(rectangle["bottom"] * image.height))
                    if bounds[2]-bounds[0] < 2 or bounds[3]-bounds[1] < 2:
                        raise ValueError("区域宽高至少为 2 像素")
                    image_path = directory / "images" / "regions" / f"{identifier}.png"
                    if image_storage._safe_path(directory, image_path) is None:
                        raise ValueError("区域截图路径必须位于当前软件目录内")
                    image_path.parent.mkdir(parents=True, exist_ok=True)
                    image.crop(bounds).save(image_path)
                    image_ref = str(image_path.relative_to(directory)).replace("\\", "/")
            knowledge.save_region(directory, {
                **candidate, "id": f"{self.page_id}.{identifier}", "app_id": app_id,
                "page_id": self.page_id, "description": description,
                "source": "manual", "geometry_role": candidate.get("geometry_role", "visual-hint-only"), "image": image_ref,
            })
            self._reassign(directory, [*regions, candidate])
            self._refresh(directory)
        return f"{self.page_id}.{identifier}"

    def save_picked_region(self, app_id: str, result: dict, *, region_id="", name="", description="") -> str:
        """Validate selection and overlap before replacing the saved reference image."""
        from PIL import Image

        selection, capture = result["selection"], result["capture"]
        location = normalize_location(selection["location"])
        if not location["Xpath"]: raise ValueError("所选区域没有可保存的 XPath")
        size = dict(width=capture["width"], height=capture["height"])
        with Image.open(capture["image"]) as image:
            if image.size != (size["width"], size["height"]): raise ValueError("区域截图尺寸不一致")
        title = name or selection["native_name"] or selection["automation_id"] or selection["control_type"]
        with image_storage._vault_lock(self.root):
            existing = self.layout_regions(app_id)
            if existing and self.page(app_id).get("window_scope",{"kind":"main"})!=capture.get("window_scope",{"kind":"main"}):
                raise ValueError("区域所属窗口已改变，请添加独立区域或区域状态，不要覆盖当前窗口布局")
            identifier = region_id.removeprefix(self.page_id + ".") if region_id else "picked-candidate"
            if region_id and not any(r["id"] == identifier for r in existing):
                raise ValueError("要修改的区域已经不存在，请重新选择")
            manual_layout.validate_regions([*[r for r in existing if r["id"] != identifier],
                dict(id=identifier, name=title, normalized_rect=selection["normalized_rect"])])
            self.save_capture(app_id, capture)
            saved = self.save_region(app_id, title, selection["normalized_rect"], region_id, description=description,
                control_record={**selection, "location": location, "picked_image_size": size})
            page = self.page(app_id)
            if not page.get("requires_layout_review") or all(
                    r.get("picked_image_size") == size for r in self.regions(app_id)):
                self.confirm_layout(app_id)
        return saved

    def apply_automatic_layout(self, app_id: str, proposal: dict) -> dict:
        """Replace untouched suggestions; keep all hand drawn or edited regions."""
        from .mcp.automatic_layout import area, exclude

        directory = self.directory(app_id)
        with image_storage._vault_lock(self.root):
            existing = self.layout_regions(app_id)
            locked = [r for r in existing if not str(r.get("layout_origin", "")).startswith("auto-")
                      or r.get("edited_by_user", False)]
            locked_ids = {r["id"] for r in locked}
            generated = []
            for region in manual_layout.validate_regions(proposal.get("regions", [])):
                parts = exclude(region["normalized_rect"], [r["normalized_rect"] for r in locked])
                for index, rectangle in enumerate(parts):
                    if area(rectangle) < .002: continue
                    identifier = region["id"] if len(parts) == 1 else knowledge.stable_id("auto-region", region["id"], index)
                    if identifier in locked_ids:
                        identifier = knowledge.stable_id("auto-region", region["id"], *(round(rectangle[k], 6)
                            for k in ("left", "top", "right", "bottom")))
                    generated.append({**region, "id": identifier, "normalized_rect": rectangle,
                                      "edited_by_user": False})
            merged = manual_layout.validate_regions([*locked, *generated])
            page = self.page(app_id)
            source = self.asset(app_id, page.get("layout_image", ""))
            if source is None: raise ValueError("请先获取软件截图")
            from PIL import Image
            prepared = []
            with Image.open(source) as image:
                for r in generated:
                    n = r["normalized_rect"]
                    pixels = tuple(round(n[k]*(image.width if k in {"left", "right"} else image.height))
                                   for k in ("left", "top", "right", "bottom"))
                    if pixels[2]-pixels[0] < 2 or pixels[3]-pixels[1] < 2: continue
                    prepared.append({**r, "image": image_storage.save_interaction_image(directory, image.crop(pixels))})
            merged = manual_layout.validate_regions([*locked, *prepared])
            kept_ids = {r["id"] for r in merged}
            archive = directory / ".easy_uiauto" / "archive" / uuid4().hex / "regions"
            if image_storage._safe_path(directory, archive) is None:
                raise ValueError("区域归档路径必须位于当前软件目录内")
            for path, record in knowledge.iter_records(directory, "region"):
                if (record.get("page_id") == self.page_id and record.get("source") == "manual"
                        and record["id"].removeprefix(self.page_id+".") not in kept_ids
                        and not record.get("classification_only")):
                    archive.mkdir(parents=True, exist_ok=True)
                    path.replace(archive / path.name)
            for r in prepared:
                knowledge.save_region(directory, {**r, "id": self.page_id+"."+r["id"], "app_id": app_id,
                    "page_id": self.page_id, "source": "manual", "geometry_role": "visual-hint-only"})
            page.update(layout_confirmed=False, layout_proposal_method=proposal.get("method", ""),
                        layout_proposal_warning=proposal.get("warning", ""))
            knowledge.save_page(directory, page)
            self._reassign(directory, merged)
            self._refresh(directory)
            return {"suggested": len(prepared), "preserved": len(locked), "warning": proposal.get("warning", "")}

    def confirm_layout(self, app_id: str) -> dict:
        directory = self.directory(app_id)
        regions = manual_layout.validate_regions(self.layout_regions(app_id))
        page = self.page(app_id)
        if not page.get("layout_image"):
            raise ValueError("请先刷新软件截图")
        page["layout_confirmed"] = True
        page["requires_layout_review"] = False
        with image_storage._vault_lock(self.root):
            knowledge.save_page(directory, page)
            self._refresh(directory)
        return {"page": page, "regions": regions}

    def learning_layout(self, app_id: str, window: dict, *, region_id: str = "") -> dict:
        page = self.page(app_id)
        if not page.get("layout_confirmed"):
            raise ValueError("请先确认区域布局")
        regions = manual_layout.validate_regions(self.layout_regions(app_id))
        if region_id:
            local = region_id.removeprefix(self.page_id+".")
            regions = [r for r in regions if r["id"] == local]
            if not regions: raise ValueError("要学习的区域已不存在")
        return {"app_id": app_id, "page_id": self.page_id, "page_name": page.get("name", "人工区域布局"),
                "window_handle": window["handle"], "window_process_id": window["process_id"],
                "window_class": window["class_name"], "image_size": page["observed_window_size"],
                "regions": regions, "restrict_regions": bool(region_id),
                "window_scope": page.get("window_scope", {"kind":"main"})}

    def _reassign(self, directory: Path, regions: list[dict]):
        used_commands = set()
        for _path, record in knowledge.iter_records(directory, "control"):
            if record.get("page_id") != self.page_id or not record.get("normalized_rect"):
                continue
            previous = record.get("region_id", "unassigned").removeprefix(self.page_id + ".")
            identifier, reason = manual_layout.assign_region(record["normalized_rect"], regions)
            record.update(region_id=identifier, region_assignment=reason, region_source="manual")
            command = record.get("command", "")
            prefix = f"{self.page_id}.{previous}."
            if command.startswith(prefix):
                command = f"{self.page_id}.{identifier}." + command[len(prefix):]
            if command and command in used_commands:
                command += "." + record["id"]
            if command:
                used_commands.add(command)
                record["command"] = command
            knowledge.save_control(directory, record)

    @staticmethod
    def _refresh(directory: Path):
        knowledge.rebuild_index(directory)
        knowledge.write_command_catalog(directory)

    def update_control(self, app_id: str, control_id: str, *, name: str,
                       description: str, aliases: list[str], location: dict | None = None,
                       match_mode: str | None = None, navigation_path: list[str] | None = None,
                       hierarchy_parent: str | None = None, entry_steps: list[dict] | None = None) -> dict:
        directory = self.directory(app_id)
        with image_storage._vault_lock(self.root):
            found = knowledge.find_control_record(directory, control_id, include_quarantine=True)
            if not found:
                raise ValueError("控件记录不存在")
            _path, record = found
            record.update(semantic_name=name.strip() or record.get("name") or record["id"],
                          description=description, aliases=aliases,
                          semantic_source="manual", semantic_status="manual",
                          semantic_confidence=1.0)
            if location is not None:
                updated = normalize_location(location)
                if match_mode is not None:
                    updated = with_match_mode(updated, match_mode)
                from .locator_images import is_image_location
                if not updated["Xpath"] and not is_image_location(updated):
                    raise ValueError("控件 LOCATION 必须保留 XPath")
                if updated != record.get("location"):
                    record.update(location=updated, status="observed",
                                  name=updated["Name"], class_name=updated["ClassName"],
                                  automation_id=updated["AutomationId"],
                                  control_type="VisualElement" if is_image_location(updated) else updated["ControlType"])
                    record["verification"] = {**record.get("verification", {}),
                                              "location": "pending", "location_detail": "Locator edited manually"}
            if navigation_path is not None:
                from .library_actions import validate_plan

                controls = [record if r["id"] == control_id else r for r in self.read(app_id)["controls"]]
                validate_plan(controls, control_id, navigation_path)
                record.update(navigation_path=navigation_path, navigation_source="manual")
            if hierarchy_parent is not None:
                from .mcp.control_hierarchy import validate_parent

                controls = [record if r["id"] == control_id else r for r in self.read(app_id)["controls"]]
                validate_parent(controls, control_id, hierarchy_parent)
                record["hierarchy_parent"] = hierarchy_parent
            if entry_steps is not None:
                from . import workflows
                if entry_steps:
                    report=workflows.readiness(dict(keyword="控件前置步骤",steps=entry_steps),self.read(app_id)["controls"])
                    if not report["ready"]:raise ValueError("控件前置步骤尚不可执行")
                record["entry_steps"]=json.loads(json.dumps(entry_steps))
            knowledge.save_control(directory, record)
            self._refresh(directory)
            return knowledge.find_control_record(directory, control_id, include_quarantine=True)[1]

    def click_plan(self, app_id: str, control_id: str) -> list[dict]:
        from .library_actions import validate_plan

        return validate_plan(self.read(app_id)["controls"], control_id)

    def save_execution(self, app_id: str, control_id: str, execution: dict) -> None:
        """Persist input outcomes without claiming a business effect or releasing quarantine."""
        directory = self.directory(app_id)
        with image_storage._vault_lock(self.root):
            found = knowledge.find_control_record(directory, control_id, include_quarantine=True)
            if found:
                record = found[1]
                record["last_click_execution"] = {**execution, "at": knowledge.utc_now()}
                knowledge.save_control(directory, record)
                try:
                    image_storage.cleanup_images(directory, dry_run=False)
                except (OSError, ValueError) as error:
                    execution["image_cleanup_error"] = str(error)
                self._refresh(directory)

    def capture_before_click(self, app_id: str, control_id: str, control, main_window,
                             *, update_reference: bool = False) -> dict:
        """Capture execution evidence; only explicit learning may update references."""
        from .mcp import control_observation, scanner

        directory = self.directory(app_id)
        with image_storage._vault_lock(self.root):
            found = knowledge.find_control_record(directory, control_id, include_quarantine=True)
            if not found:
                raise ValueError("控件记录已删除，未执行点击")
            observation, crop = control_observation.capture(directory, control, main_window, phase="before-action")
            observation.update(purpose="execution-evidence",
                               control_image=image_storage.save_interaction_image(directory, crop))
            if update_reference:
                record = found[1]
                path, variants, digest = scanner._save_control_images(directory, control_id, crop, record)
                reference = {**observation, "purpose": "learning-reference"}
                record.update(image=str(path.relative_to(directory)).replace("\\", "/"),
                              image_variants=variants, image_sha256=digest, visual_observation=reference)
                knowledge.save_control(directory, record)
            return observation

    def archive(self, app_id: str, kind: str, identifier: str) -> Path:
        return self.archive_many([dict(app_id=app_id,kind=kind,id=identifier,page_id=self.page_id)])[0]

    def archive_many(self, items: list[dict]) -> list[Path]:
        """Archive explicit selections; keep unselected descendants and refresh once per app."""
        if not items:return []
        with image_storage._vault_lock(self.root):
            batch=uuid4().hex;plan=[];seen=set();apps={}
            for item in items:
                app_id,kind,identifier=item["app_id"],item["kind"],item["id"]
                if kind not in {"control","region","reference"}:raise ValueError("只能删除区域、控件或参考资料")
                key=app_id,kind,identifier
                if key in seen:continue
                seen.add(key)
                if app_id not in apps:
                    directory=self.directory(app_id)
                    records=knowledge.iter_records(directory)
                    apps[app_id]=(directory,records)
                directory,records=apps[app_id]
                matches=[(p,r) for p,r in records if r.get("kind")==kind and r.get("id")==identifier]
                if len(matches)!=1:raise ValueError("所选记录已不存在，请刷新知识库")
                path,record=matches[0]
                if kind=="region" and not dynamic_regions.is_dynamic(record) and (
                        record.get("source")!="manual" or record.get("classification_only")
                        or record.get("page_id")!=item.get("page_id",record.get("page_id"))):
                    raise ValueError("历史自动区域和待归档节点仅供查看；可删除其中的控件记录")
                destination=directory/".easy_uiauto/archive"/batch/path.parent.name/path.name
                if image_storage._safe_path(directory,path) is None or image_storage._safe_path(directory,destination) is None:
                    raise ValueError("归档路径必须位于当前软件目录内")
                plan.append((app_id,kind,path,record,destination))

            # Validate the entire selection before moving a file. These small text
            # snapshots also restore assignments/indexes if a later write fails.
            backups={}
            for directory,records in apps.values():
                for path in [*(p for p,_r in records),directory/".easy_uiauto/index.json",directory/"operations/UI-CLI.md"]:
                    if image_storage._safe_path(directory,path) is None:raise ValueError("知识库记录不能为链接")
                    backups[path]=path.read_bytes() if path.is_file() else None
            moved=[]
            try:
                for _app,_kind,path,_record,destination in plan:
                    destination.parent.mkdir(parents=True,exist_ok=True)
                    path.replace(destination);moved.append(destination)
                for app_id,(directory,_records) in apps.items():
                    removed={r["id"]:r for app,kind,_p,r,_d in plan if app==app_id and kind=="region"}
                    for record in removed.values():
                        if not dynamic_regions.is_dynamic(record):continue
                        source_id=record.get("source_control_id")
                        found=knowledge.find_control_record(directory,source_id) if source_id else None
                        if found:
                            found[1]["region_behavior"]="control";knowledge.save_control(directory,found[1])
                    for _path,child in knowledge.iter_records(directory,"region"):
                        parent=child.get("parent_region_id");visited=set()
                        while parent in removed and parent not in visited:
                            visited.add(parent);record=removed[parent]
                            child["parent_region_id"]=record.get("parent_region_id","")
                            child["parent_control_id"]=record.get("parent_control_id","")
                            parent=child["parent_region_id"]
                        if visited:knowledge.save_region(directory,child)
                    pages={r["page_id"] for r in removed.values() if not dynamic_regions.is_dynamic(r)}
                    for page_id in pages:
                        scoped=self.scoped(page_id);scoped._reassign(directory,scoped.layout_regions(app_id))
                    self._refresh(directory)
            except Exception as error:
                failures=[]
                for path,content in backups.items():
                    try:
                        if content is None:path.unlink(missing_ok=True)
                        else:path.write_bytes(content)
                    except OSError as restore_error:failures.append(str(restore_error))
                if not failures:
                    for path in moved:
                        try:path.unlink()
                        except OSError as restore_error:failures.append(str(restore_error))
                detail=("原记录已恢复" if not failures else "恢复未完整完成，请保留归档目录："+str(self.root))
                raise RuntimeError(f"批量删除失败，{detail}：{error}") from error
            return [destination for _app,_kind,_path,_record,destination in plan]

    def export(self, app_id: str, destination: str | Path) -> Path:
        payload = self.read(app_id)
        # Image references remain relative to the application directory.
        path = Path(destination)
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return path

    @staticmethod
    def python_code(record: dict) -> str:
        location = record.get("location")
        if not location:
            return "# 此记录只有图像观察，没有可回放的 UIA LOCATION。\n"
        value = json.dumps(location, ensure_ascii=False, indent=2)
        return ("import json\nfrom easy_uiauto import find_control\n\n"
                f"LOCATION = json.loads({value!r})\ncontrol = find_control(LOCATION)\n")
