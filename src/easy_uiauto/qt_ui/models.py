"""Knowledge navigation and step tables keep stable record references."""
from __future__ import annotations

from PySide6.QtCore import Qt, QSortFilterProxyModel
from PySide6.QtGui import QStandardItem, QStandardItemModel

from ..mcp import dynamic_regions, knowledge
from .icons import icon, status_icon_name

ROLE = Qt.ItemDataRole.UserRole + 1
NODE_ICONS = {"vault":"library", "application":"window", "usage-root":"book-open",
              "usage":"documents", "references-root":"documents", "reference":"documents",
              "controls-root":"controls", "area":"region", "region":"region", "control":"control",
              "actions-root":"workflow", "workflow":"workflow"}


def control_status(record):
    status = record.get("status")
    verification = record.get("verification", {})
    if knowledge.runtime_quarantined(record):return "已隔离"
    if status == "quarantined":
        if verification.get("image")=="passed" or record.get("visual_fallback_ready"):
            return "功能待确认" if verification.get("location")=="passed" else "图片定位待验证"
        return "定位待复核"
    if record.get("control_type") == "VisualElement":
        return "图片定位已验证" if status=='verified' else "图片定位待验证"
    return {"verified": "定位已验证", "suspect": "待复核"}.get(status, "已采集")


class KnowledgeModel(QStandardItemModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHorizontalHeaderLabels(["知识节点", "状态"])
        self.nodes = {}

    def add(self, parent, key, text, app_id, kind, record, status=""):
        item = QStandardItem(text); item.setEditable(False)
        name = NODE_ICONS.get(kind, "")
        if kind == "region" and record.get("state_name"):name = "layers"
        item.setIcon(icon(name))
        item.setData(dict(key=key, app_id=app_id, kind=kind, record=record), ROLE)
        flag = QStandardItem(status); flag.setEditable(False)
        flag.setIcon(icon(status_icon_name(status)))
        parent.appendRow([item, flag]); self.nodes[key] = item
        return item

    def load(self, vault_name, snapshots, windows):
        self.clear(); self.setHorizontalHeaderLabels(["知识节点", "状态"]); self.nodes.clear()
        root = self.add(self.invisibleRootItem(), "vault", vault_name, "", "vault", {})
        for app_id, index in snapshots.items():
            application = index["application"]
            app = self.add(root, "app:"+app_id, application.get("name", app_id), app_id, "application", application,
                           "已连接" if app_id in windows else "未连接")
            usage = self.add(app, "usage-root:"+app_id, "使用知识", app_id, "usage-root", {})
            references = self.add(usage, "references-root:"+app_id, "参考资料", app_id, "references-root", {})
            for r in index.get("references", []):
                self.add(references, f"reference:{app_id}:{r['id']}", r["name"], app_id, "reference", r,
                         "已整理" if r.get("enabled", True) else "已停用")
            for r in index["usage"]:
                self.add(usage, f"usage:{app_id}:{r['document_path']}", r["name"], app_id, "usage", r)
            controls_root = self.add(app, "controls-root:"+app_id, "控件库", app_id, "controls-root", {})
            pages = {p["id"]:p for p in index["pages"]}
            region_items, control_items = {}, {}
            static = [r for r in index["regions"] if not dynamic_regions.is_dynamic(r)]
            groups = {}
            for r in static:
                state = pages.get(r.get("page_id"), {}).get("region_state", {})
                groups.setdefault(state.get("area_id") or r["id"], []).append(r)
            for area_id, regions in groups.items():
                base = next((r for r in regions if r["id"] == area_id), regions[0])
                parent = controls_root
                if len(regions) > 1:
                    parent = self.add(controls_root, f"area:{app_id}:{area_id}", base.get("name", "区域"), app_id, "area", base)
                for r in regions:
                    page = pages.get(r.get("page_id"), {})
                    state = page.get("region_state", {})
                    label = (r.get("state_name") or state.get("name") or "默认") if len(regions)>1 else r.get("name",r["id"])
                    region_items[r["id"]] = self.add(parent, f"region:{app_id}:{r['id']}", label, app_id, "region", r,
                        "待确认" if not page.get("layout_confirmed") else "有前置" if r.get("entry_steps") or state.get("steps") else "")
            graph = dynamic_regions.build_tree(index["controls"], index["regions"])
            pending = {}
            for r in index["regions"]:
                if dynamic_regions.is_dynamic(r):
                    pending["region:"+r["id"]] = ("region", r)
            for r in index["controls"]:
                if r["id"] not in graph.get("containers", {}): pending["control:"+r["id"]] = ("control", r)
            placed = {"region:"+identifier:item for identifier,item in region_items.items()}
            while pending:
                ready = [n for n in pending if not graph["parents"].get(n) or graph["parents"].get(n) in placed]
                if not ready: ready = [next(iter(pending))]
                for node in ready:
                    kind, r = pending.pop(node); parent = placed.get(graph["parents"].get(node), controls_root)
                    if kind == "control":
                        label = r.get("semantic_name") or r.get("name") or r["id"]
                        status = control_status(r)
                    else: label, status = r.get("name", r["id"]), "展开区域"
                    placed[node] = self.add(parent, f"{kind}:{app_id}:{r['id']}", label, app_id, kind, r, status)
            actions = self.add(app, "actions-root:"+app_id, "组合动作", app_id, "actions-root", {})
            for flow in index["workflows"]:
                self.add(actions, f"workflow:{app_id}:{flow['id']}", flow["keyword"], app_id, "workflow", flow,
                         "可运行" if flow.get("status") == "ready" else "待完善")


class KnowledgeFilter(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setRecursiveFilteringEnabled(True)
        self.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.setFilterKeyColumn(0)
