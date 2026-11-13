"""Knowledge navigation and step tables keep stable record references."""
from __future__ import annotations

from collections import defaultdict, deque

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
    if record.get("availability", {}).get("state") == "unobserved":return "更新后待定位"
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
        self.rows = {}

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

    @staticmethod
    def describe(vault_name, snapshots, windows):
        """Build the hierarchy off the GUI thread; descriptions contain no Qt objects."""
        rows = []
        def add(parent, key, text, app_id, kind, record, status=""):
            rows.append((parent, key, text, app_id, kind, record, status))
            return key
        root = add(None, "vault", vault_name, "", "vault", {})
        for app_id, index in snapshots.items():
            application = index["application"]
            app = add(root, "app:"+app_id, application.get("name", app_id), app_id, "application", application,
                           "已连接" if app_id in windows else "未连接")
            usage = add(app, "usage-root:"+app_id, "使用知识", app_id, "usage-root", {})
            references = add(usage, "references-root:"+app_id, "参考资料", app_id, "references-root", {})
            for r in index.get("references", []):
                add(references, f"reference:{app_id}:{r['id']}", r["name"], app_id, "reference", r,
                         "已整理" if r.get("enabled", True) else "已停用")
            for r in index["usage"]:
                add(usage, f"usage:{app_id}:{r['document_path']}", r["name"], app_id, "usage", r)
            controls_root = add(app, "controls-root:"+app_id, "控件库", app_id, "controls-root", {})
            pages = {p["id"]:p for p in index["pages"]}
            region_items = {}
            static = [r for r in index["regions"] if not dynamic_regions.is_dynamic(r)]
            groups = {}
            for r in static:
                state = pages.get(r.get("page_id"), {}).get("region_state", {})
                groups.setdefault(state.get("area_id") or r["id"], []).append(r)
            for area_id, regions in groups.items():
                base = next((r for r in regions if r["id"] == area_id), regions[0])
                parent = controls_root
                if len(regions) > 1:
                    parent = add(controls_root, f"area:{app_id}:{area_id}", base.get("name", "区域"), app_id, "area", base)
                for r in regions:
                    page = pages.get(r.get("page_id"), {})
                    state = page.get("region_state", {})
                    label = (r.get("state_name") or state.get("name") or "默认") if len(regions)>1 else r.get("name",r["id"])
                    region_items[r["id"]] = add(parent, f"region:{app_id}:{r['id']}", label, app_id, "region", r,
                        "待确认" if not page.get("layout_confirmed") else "有前置" if r.get("entry_steps") or state.get("steps") else "")
            graph = dynamic_regions.build_tree(index["controls"], index["regions"])
            pending = {}
            for r in index["regions"]:
                if dynamic_regions.is_dynamic(r):
                    pending["region:"+r["id"]] = ("region", r)
            for r in index["controls"]:
                if r["id"] not in graph.get("containers", {}): pending["control:"+r["id"]] = ("control", r)
            placed = {"region:"+identifier:item for identifier,item in region_items.items()}
            children = defaultdict(list)
            ready = deque()
            for node in pending:
                parent = graph["parents"].get(node)
                if parent not in pending:ready.append(node)
                else:children[parent].append(node)
            while pending:
                if not ready:ready.append(next(iter(pending)))
                node = ready.popleft()
                if node not in pending:continue
                kind, r = pending.pop(node); parent = placed.get(graph["parents"].get(node), controls_root)
                if kind == "control":
                    label = r.get("semantic_name") or r.get("name") or r["id"]
                    status = control_status(r)
                else: label, status = r.get("name", r["id"]), "展开区域"
                placed[node] = add(parent, f"{kind}:{app_id}:{r['id']}", label, app_id, kind, r, status)
                ready.extend(children[node])
            actions = add(app, "actions-root:"+app_id, "组合动作", app_id, "actions-root", {})
            for flow in index["workflows"]:
                add(actions, f"workflow:{app_id}:{flow['id']}", flow["keyword"], app_id, "workflow", flow,
                         "更新后待验证" if flow.get("update_validation",{}).get("status") in {"pending","failed"} else
                         "可运行" if flow.get("status") == "ready" else "待完善")
        return rows

    def load(self, vault_name, snapshots, windows):
        self.sync(self.describe(vault_name, snapshots, windows))

    def sync(self, rows, app_ids=None):
        """Insert/update/move changed rows while retaining unaffected item identities."""
        wanted = {row[1] for row in rows}
        obsolete = {key for key, row in self.rows.items()
                    if key not in wanted and (app_ids is None or row[3] in app_ids)}
        for row in rows:
            parent_key, key, text, app_id, kind, record, status = row
            parent = self.nodes[parent_key] if parent_key else self.invisibleRootItem()
            item = self.nodes.get(key)
            if item is None:
                item = self.add(parent, key, text, app_id, kind, record, status)
            else:
                old_parent = item.parent() or self.invisibleRootItem()
                if old_parent is not parent:
                    parent.appendRow(old_parent.takeRow(item.row()))
                old = self.rows[key]
                if old[2] != text:item.setText(text)
                if old[4:6] != row[4:6]:
                    item.setData(dict(key=key, app_id=app_id, kind=kind, record=record), ROLE)
                    name = "layers" if kind == "region" and record.get("state_name") else NODE_ICONS.get(kind, "")
                    item.setIcon(icon(name))
                if old[6] != status:
                    flag = parent.child(item.row(), 1)
                    flag.setText(status); flag.setIcon(icon(status_icon_name(status)))
            self.rows[key] = row
        # Move surviving children first, then remove obsolete branches bottom-up.
        def depth(key):
            value = self.nodes[key]; result = 0
            while value.parent() is not None:result += 1; value = value.parent()
            return result
        for key in sorted(obsolete, key=depth, reverse=True):
            item = self.nodes.pop(key); self.rows.pop(key)
            (item.parent() or self.invisibleRootItem()).removeRow(item.row())
        positions = {}
        for parent_key, key, *_ in rows:
            parent = self.nodes[parent_key] if parent_key else self.invisibleRootItem()
            position = positions.get(parent_key, 0); positions[parent_key] = position + 1
            # A partial app refresh must not move that app ahead of other apps.
            if parent_key == "vault" and app_ids is not None:continue
            item = self.nodes[key]
            if item.row() != position:parent.insertRow(position, parent.takeRow(item.row()))


class KnowledgeFilter(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setRecursiveFilteringEnabled(True)
        self.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.setFilterKeyColumn(0)
