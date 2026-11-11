"""Conservative, editable region proposals from UIA containers and pixel dividers."""

from __future__ import annotations

from . import knowledge, manual_layout

KEYS = ("left", "top", "right", "bottom")
ROLES = {
    "TitleBarControl": (0, "标题栏"), "MenuBarControl": (1, "菜单栏"),
    "ToolBarControl": (2, "工具栏"), "StatusBarControl": (3, "状态栏"),
    "DocumentControl": (4, "文档区"), "TreeControl": (5, "导航区"),
    "TableControl": (5, "表格区"), "DataGridControl": (5, "表格区"),
    "ListControl": (6, "列表区"), "TabControl": (7, "选项卡区域"),
    "PaneControl": (8, "面板"), "GroupControl": (9, "分组面板"),
}


def area(r):
    return max(0, r["right"]-r["left"])*max(0, r["bottom"]-r["top"])


def subtract(rect, cut):
    l, t = max(rect["left"], cut["left"]), max(rect["top"], cut["top"])
    r, b = min(rect["right"], cut["right"]), min(rect["bottom"], cut["bottom"])
    if l >= r or t >= b:
        return [rect]
    values = [(rect["left"], rect["top"], rect["right"], t),
              (rect["left"], b, rect["right"], rect["bottom"]),
              (rect["left"], t, l, b), (r, t, rect["right"], b)]
    return [dict(zip(KEYS, v)) for v in values if v[2] > v[0] and v[3] > v[1]]


def exclude(rect, cuts):
    pieces = [rect]
    for cut in cuts:
        pieces = [part for piece in pieces for part in subtract(piece, cut)]
    return pieces


def _flatten(tree, width, height, origin):
    pending, result = [(tree, 0)], []
    while pending:
        node, depth = pending.pop()
        if not isinstance(node, dict): continue
        pending.extend((c, depth+1) for c in reversed(node.get("children", [])))
        raw = node.get("bounds")
        if not raw or depth == 0: continue
        try:
            l, t, r, b = (float(raw[k])-origin[0 if k in {"left", "right"} else 1] for k in KEYS)
            rect = dict(left=max(0, l), top=max(0, t), right=min(width, r), bottom=min(height, b))
            if area(rect) < 4: continue
        except (KeyError, TypeError, ValueError):
            continue
        result.append({"rect": rect, "type": node.get("ControlType"), "name": str(node.get("Name") or ""),
                       "children": len(node.get("children", [])), "depth": depth})
    return result


def _pixel_parts(image, rect, budget=6):
    """Only split on long straight pixel boundaries, never OCR word boxes."""
    import numpy as np

    gray = np.asarray(image.convert("L"), dtype="int16")
    pending, result = [rect], []
    while pending:
        current = pending.pop(0)
        l, t, r, b = (round(current[k]) for k in KEYS)
        patch = gray[t:b, l:r]
        candidates = []
        for axis in (0, 1):
            span = patch.shape[axis]
            length = patch.shape[1-axis]
            minimum = max(22 if axis == 0 else 100, round((image.height if axis == 0 else image.width)*.035))
            if span < minimum*2 or length < 120: continue
            strength = (abs(np.diff(patch, axis=axis)) >= 14).mean(axis=1-axis)
            indexes = np.flatnonzero(strength >= .72)
            for index in indexes:
                offset = int(index)+1
                if minimum <= offset <= span-minimum:
                    balance = min(offset, span-offset)/span
                    candidates.append((float(strength[index])*.2+balance, axis, offset))
        if not candidates or len(result)+len(pending)+1 >= budget:
            result.append(current); continue
        _score, axis, offset = max(candidates)
        if axis == 0:
            parts = [dict(left=l, top=t, right=r, bottom=t+offset), dict(left=l, top=t+offset, right=r, bottom=b)]
        else:
            parts = [dict(left=l, top=t, right=l+offset, bottom=b), dict(left=l+offset, top=t, right=r, bottom=b)]
        pending.extend(parts)
    return result


def propose(image, tree=None, *, origin=(0, 0), inspection=None):
    width, height = image.size
    full = dict(left=0, top=0, right=width, bottom=height)
    nodes = _flatten(tree, width, height, origin)
    candidates = []
    for node in nodes:
        kind, rect = node["type"], node["rect"]
        if kind not in ROLES: continue
        priority, role = ROLES[kind]
        w, h = rect["right"]-rect["left"], rect["bottom"]-rect["top"]
        if w < 24 or h < 12 or area(rect) > width*height*.94: continue
        if kind in {"PaneControl", "GroupControl"} and (
                area(rect) < width*height*.025 or w < 100 or h < 50
                or node["children"] < 2 and not node["name"].strip()):
            continue
        label = node["name"].strip()[:40]
        if kind in {"TitleBarControl", "MenuBarControl", "ToolBarControl", "StatusBarControl"}:
            label = role
            if w > width*.7:
                rect = {**rect, "left": 0, "right": width}
        candidates.append((priority, area(rect), label or role, rect, kind))

    # Several aligned buttons form one toolbar; a button never becomes a region.
    rows = []
    for node in nodes:
        r = node["rect"]
        if node["type"] != "ButtonControl" or r["top"] > height*.35 or r["bottom"]-r["top"] > 80: continue
        cy = (r["top"]+r["bottom"])/2
        row = next((items for items in rows if abs(items[0]-cy) <= 6), None)
        if row is None: rows.append([cy, [r]])
        else: row[1].append(r)
    for _cy, row in rows:
        if len(row) < 3: continue
        left, right = min(r["left"] for r in row), max(r["right"] for r in row)
        if right-left < width*.25: continue
        rect = dict(left=0, top=max(0, min(r["top"] for r in row)-8), right=width,
                    bottom=min(height, max(r["bottom"] for r in row)+8))
        candidates.append((2.5, area(rect), "工具栏", rect, "aligned-buttons"))

    chosen = []
    for _priority, _size, name, rect, evidence in sorted(candidates, key=lambda c: (c[0], c[1], c[2])):
        pieces = exclude(rect, [r["rect"] for r in chosen])
        if not pieces: continue
        largest = max(pieces, key=area)
        if area(largest) < area(rect)*.55: continue
        chosen.append(dict(name=name, rect=largest, evidence=evidence, origin="auto-uia"))
        if len(chosen) >= 10: break
    # A small gap above an aligned command row belongs to the same toolbar,
    # rather than becoming a new empty region.
    for item in chosen:
        rect = item["rect"]
        if item["evidence"] != "aligned-buttons": continue
        upper = [r["rect"]["bottom"] for r in chosen if r is not item and r["rect"]["bottom"] <= rect["top"]]
        if upper and 0 <= rect["top"]-max(upper) <= min(60, height*.1):
            gap = {**rect, "top": max(upper), "bottom": rect["top"]}
            if all(not area_overlap(gap, other["rect"]) for other in chosen if other is not item):
                item["rect"] = {**rect, "top": max(upper)}
    native_structure = len(chosen) >= 2
    remainder = exclude(full, [r["rect"] for r in chosen])
    for rect in sorted(remainder, key=area, reverse=True):
        if area(rect) < width*height*.025 or rect["right"]-rect["left"] < 60 or rect["bottom"]-rect["top"] < 24: continue
        parts = [rect] if native_structure else _pixel_parts(image, rect, budget=min(4, max(1, 8-len(chosen))))
        for part in parts:
            if area(part) < width*height*.015: continue
            name = "内容区" if area(part) >= width*height*.2 else "界面区域"
            chosen.append(dict(name=name, rect=part,
                               evidence="uncovered-content" if native_structure else "long-pixel-dividers", origin="auto-pixels"))
        if len(chosen) >= 12: break
    if not chosen:
        chosen = [dict(name="内容区", rect=full, evidence="whole-window-fallback", origin="auto-pixels")]
    chosen.sort(key=lambda r: (r["rect"]["top"], r["rect"]["left"], r["name"]))
    counts, regions = {}, []
    for item in chosen:
        name = item["name"]
        counts[name] = counts.get(name, 0)+1
        if counts[name] > 1: name += f" {counts[name]}"
        n = {k: item["rect"][k]/(width if k in {"left", "right"} else height) for k in KEYS}
        regions.append(dict(id=knowledge.stable_id("auto-region", item["evidence"], *(round(n[k], 4) for k in KEYS)),
                            name=name, normalized_rect=n, layout_origin=item["origin"],
                            region_evidence=item["evidence"], edited_by_user=False))
    manual_layout.validate_regions(regions)
    return dict(regions=regions, inspection=inspection or {}, uia_nodes=len(nodes),
                method="uia-containers-and-pixel-dividers" if nodes else "pixel-dividers")


def area_overlap(first, second):
    return max(0, min(first["right"], second["right"])-max(first["left"], second["left"]))*max(
        0, min(first["bottom"], second["bottom"])-max(first["top"], second["top"]))
