"""Render and validate the hand-drawn SVG set using the GUI's Qt SVG renderer.

Run from the repository with the GUI dependencies installed:
    python scripts/preview_gui_icons.py
SVG files are the editable source; the manifest provides names and colors.
"""
from __future__ import annotations

import html
import json
import math
import os
from pathlib import Path
import re
import xml.etree.ElementTree as ET

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QByteArray, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QFontDatabase, QImage, QPainter
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import QApplication


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "src/easy_uiauto/resources/icons"
OUTPUT = ROOT / "docs/assets/gui-icons"
GROUPS = {"navigation": "导航与内容", "common": "通用操作", "learning": "学习与流程", "status": "状态提示"}


def recolor(svg, color):
    return re.sub(r' color="#[0-9A-Fa-f]{6}"', f' color="{color}"', svg, count=1)


def raster(svg, size):
    renderer = QSvgRenderer(QByteArray(svg.encode("utf-8")))
    assert renderer.isValid(), "Qt rejected SVG"
    image = QImage(size, size, QImage.Format.Format_ARGB32)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    renderer.render(painter)
    painter.end()
    return image


def validate(icons):
    allowed = {"svg", "title", "path", "rect", "circle", "ellipse", "g"}
    sizes = (16, 20, 24, 32, 48)
    for icon in icons:
        svg = icon["svg"]
        tree = ET.fromstring(svg)
        assert tree.get("viewBox") == "0 0 24 24", icon["id"]
        assert tree.get("stroke-width") == "1.75", icon["id"]
        assert tree.get("color") == icon["color"], icon["id"]
        assert tree.get("stroke-linecap") == tree.get("stroke-linejoin") == "round", icon["id"]
        for element in tree.iter():
            assert element.tag.rsplit("}", 1)[-1] in allowed, icon["id"]
            assert not any("href" in key or key.startswith("on") or "url(" in value
                           for key, value in element.attrib.items()), icon["id"]
        for color in (icon["color"], "#9CA3AF", "#FFFFFF"):
            for size in sizes:
                image = raster(recolor(svg, color), size)
                colors = [image.pixelColor(x, y) for x in range(size) for y in range(size)
                          if image.pixelColor(x, y).alpha() > 0]
                assert colors, icon["id"]
                expected = QColor(color)
                # Qt composites anti-aliased edges in premultiplied color space.
                # Ignore faint fringe pixels and allow its integer roundoff.
                assert all(max(abs(c.red()-expected.red()), abs(c.green()-expected.green()),
                               abs(c.blue()-expected.blue())) <= 2
                           for c in colors if c.alpha() >= 128), (icon["id"], color, size)
                assert all(image.pixelColor(x, y).alpha() == 0 for x, y in
                           [(n, edge) for n in range(size) for edge in (0, size-1)] +
                           [(edge, n) for n in range(size) for edge in (0, size-1)]), (icon["id"], "clipped")
        # Check that the source retains two full design units of padding.
        large = raster(svg, 96)
        assert all(large.pixelColor(x, y).alpha() == 0 for x in range(96) for y in range(96)
                   if x < 8 or y < 8 or x >= 88 or y >= 88), (icon["id"], "padding")


def draw_text(painter, x, y, value, size=11, color="#475569", bold=False):
    painter.setFont(QFont("Microsoft YaHei UI", size, QFont.Weight.DemiBold if bold else QFont.Weight.Normal))
    painter.setPen(QColor(color))
    painter.drawText(x, y, value)


def contact_sheet(icons):
    width, margin, columns, gap, card_h = 1280, 36, 5, 12, 115
    card_w = (width - 2 * margin - gap * (columns - 1)) // columns
    height = 210 + sum(48 + math.ceil(sum(i["group"] == group for i in icons) / columns) * (card_h + gap)
                       for group in GROUPS)
    sheet = QImage(width, height, QImage.Format.Format_ARGB32)
    sheet.fill(QColor("#F4F7FB"))
    painter = QPainter(sheet)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    draw_text(painter, margin, 48, "easy-uiauto · GUI 图标", 24, "#0F172A", True)
    draw_text(painter, margin, 80, "49 类语义 / 50 个 SVG   ·   24 × 24 网格   ·   1.75 线宽   ·   圆角线性", 11)
    draw_text(painter, margin, 106, "每张卡片：32 px 主图，右侧依次为 16 / 20 / 24 px 原色、24 px 禁用、24 px 选中。", 10)
    x = margin
    for label, color in (("功能", "#2563EB"), ("工具", "#475569"), ("智能", "#7C3AED"),
                         ("成功", "#15803D"), ("提醒", "#B45309"), ("危险", "#DC2626")):
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(color))
        painter.drawRoundedRect(QRectF(x, 126, 12, 12), 3, 3)
        draw_text(painter, x+20, 137, label + " " + color, 9)
        x += 170
    y = 166
    for group, label in GROUPS.items():
        members = [icon for icon in icons if icon["group"] == group]
        draw_text(painter, margin, y+24, f"{label}  /  {len(members):02d}", 13, "#0F172A", True)
        y += 44
        for index, icon in enumerate(members):
            x = margin + (index % columns) * (card_w + gap)
            cy = y + (index // columns) * (card_h + gap)
            painter.setPen(QColor("#E2E8F0"))
            painter.setBrush(QColor("#FFFFFF"))
            painter.drawRoundedRect(QRectF(x, cy, card_w, card_h), 9, 9)
            painter.drawImage(x+14, cy+15, raster(icon["svg"], 32))
            for offset, size in ((62, 16), (88, 20), (118, 24)):
                painter.drawImage(x+offset, cy+31-size//2, raster(icon["svg"], size))
            painter.drawImage(x+150, cy+19, raster(recolor(icon["svg"], "#9CA3AF"), 24))
            painter.setBrush(QColor("#1E293B"))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(QRectF(x+182, cy+15, 32, 32), 5, 5)
            painter.drawImage(x+186, cy+19, raster(recolor(icon["svg"], "#FFFFFF"), 24))
            draw_text(painter, x+14, cy+74, icon["label"], 11, "#0F172A", True)
            draw_text(painter, x+14, cy+97, icon["id"], 9)
        y += math.ceil(len(members)/columns) * (card_h+gap) + 4
    painter.end()
    assert sheet.save(str(OUTPUT / "contact-sheet.png"))


def html_preview(icons):
    cards = []
    for icon in icons:
        label, identifier = html.escape(icon["label"]), html.escape(icon["id"])
        cards.append(f'''<article class="card" data-group="{icon['group']}" data-search="{label} {identifier}" style="--accent:{icon['color']}">
<div class="stage">{icon['svg']}</div><div class="label">{label}</div>
<div class="identifier">{identifier}</div><div class="meta"><span>{icon['color']}</span>
<button type="button" class="download" data-id="{identifier}">下载 SVG</button></div></article>''')
    page = '''<!doctype html>
<html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>easy-uiauto · SVG 图标预览</title>
<style>
*{box-sizing:border-box}body{margin:0;background:#f4f7fb;color:#0f172a;font:14px/1.6 "Microsoft YaHei UI",system-ui,sans-serif}
main{max-width:1350px;margin:auto;padding:40px 28px}h1{font-size:30px;letter-spacing:-.5px;margin:0 0 8px}.intro{margin:0 0 24px;color:#475569}
.toolbar{position:sticky;top:0;z-index:1;background:#f4f7fb;display:flex;flex-wrap:wrap;align-items:center;gap:12px;padding:16px 0;border-bottom:1px solid #e2e8f0}
input[type=search],select,button{font:inherit;border:1px solid #cbd5e1;border-radius:6px;background:white;color:#334155;padding:6px 10px}
input[type=search]{min-width:200px}label{display:flex;gap:8px;align-items:center}button{cursor:pointer}button:hover{background:#eff6ff;border-color:#2563eb}
:focus-visible{outline:2px solid #2563eb;outline-offset:3px}.counter{margin-left:auto;color:#64748b}.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:14px;margin-top:20px}
.card{border:1px solid #e2e8f0;border-radius:10px;background:white;padding:16px;min-width:0}.card[hidden]{display:none}.stage{height:85px;display:flex;align-items:center;justify-content:center;background:#f8fafc;border-radius:6px;margin-bottom:14px}
.stage svg{width:var(--size,32px);height:var(--size,32px);flex:none;color:var(--accent)}.label{font-weight:600}.identifier{color:#64748b;font-size:12px;margin:2px 0 12px}
.meta{display:flex;justify-content:space-between;align-items:center;gap:6px;font-size:11px;color:#64748b}.download{font-size:12px;padding:3px 6px}
body[data-mode=disabled] .stage svg{color:#9ca3af}body[data-mode=selected] .stage{background:#1e293b}body[data-mode=selected] .stage svg{color:white}
.swatches{display:flex;flex-wrap:wrap;gap:16px;margin-top:18px;color:#475569;font-size:12px}.swatch{display:inline-flex;align-items:center;gap:6px}.dot{width:10px;height:10px;border-radius:3px;background:var(--color)}
.note{color:#64748b;font-size:12px;margin-top:24px}.empty{padding:30px;text-align:center;color:#64748b}
@media(max-width:600px){main{padding:24px 16px}h1{font-size:24px}.counter{margin-left:0}.grid{grid-template-columns:repeat(2,minmax(0,1fr))}.card{padding:12px}.meta{flex-wrap:wrap}}
</style><main><h1>easy-uiauto · SVG 图标</h1>
<p class="intro">49 类语义，50 个矢量文件。统一 24 × 24 网格、1.75 线宽与圆角端点，透明背景。</p>
<div class="swatches">SWATCHES</div>
<div class="toolbar"><input id="search" type="search" aria-label="搜索图标" placeholder="搜索中文名称 / 英文名">
<label>分类<select id="group"><option value="all">全部</option><option value="navigation">导航与内容</option><option value="common">通用操作</option><option value="learning">学习与流程</option><option value="status">状态提示</option></select></label>
<label>尺寸<select id="size"><option>16</option><option>18</option><option>20</option><option>24</option><option selected>32</option><option>48</option></select>px</label>
<label>外观<select id="mode"><option value="normal">原色</option><option value="disabled">禁用灰</option><option value="selected">选中白</option></select></label>
<span class="counter" id="counter" aria-live="polite">50 / 50</span></div>
<div class="grid">CARDS</div><p class="empty" id="empty" hidden>没有匹配的图标。</p>
<p class="note">下载 SVG 保留设计原色。外观切换仅用于预览；“处理中”在本页为静态形状，接入 GUI 时可旋转。</p></main>
<script>
const cards=[...document.querySelectorAll('.card')], search=document.getElementById('search'), group=document.getElementById('group');
function filter(){let count=0;const query=search.value.trim().toLowerCase();for(const card of cards){card.hidden=!(card.dataset.search.toLowerCase().includes(query)&&(group.value==='all'||card.dataset.group===group.value));if(!card.hidden)count++}document.getElementById('counter').textContent=`${count} / ${cards.length}`;document.getElementById('empty').hidden=count!==0}
search.addEventListener('input',filter);group.addEventListener('change',filter);
document.getElementById('size').addEventListener('change',event=>document.documentElement.style.setProperty('--size',event.target.value+'px'));
document.getElementById('mode').addEventListener('change',event=>document.body.dataset.mode=event.target.value);
document.querySelectorAll('.download').forEach(button=>button.addEventListener('click',()=>{const svg=button.closest('.card').querySelector('svg').cloneNode(true);const data=new XMLSerializer().serializeToString(svg);const url=URL.createObjectURL(new Blob([data],{type:'image/svg+xml;charset=utf-8'}));const link=document.createElement('a');link.href=url;link.download=button.dataset.id+'.svg';document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),1000)}));
</script></html>'''
    swatches = ''.join(f'<span class="swatch"><span class="dot" style="--color:{color}"></span>{name} {color}</span>'
                       for name, color in (("功能蓝", "#2563EB"), ("中性灰", "#475569"), ("智能紫", "#7C3AED"),
                                           ("成功绿", "#15803D"), ("提醒橙", "#B45309"), ("危险红", "#DC2626")))
    (ROOT / "docs/gui-icons-preview.html").write_text(page.replace("CARDS", '\n'.join(cards)).replace("SWATCHES", swatches), encoding="utf-8")


def main():
    app = QApplication.instance() or QApplication([])
    for name in ("msyh.ttc", "msyhbd.ttc", "segoeui.ttf"):
        path = Path("C:/Windows/Fonts") / name
        if path.exists():
            QFontDatabase.addApplicationFont(str(path))
    manifest = json.loads((SOURCE / "manifest.json").read_text(encoding="utf-8"))
    icons = manifest["icons"]
    assert len(icons) == manifest["assetCount"] == 50
    assert len({icon["concept"] for icon in icons}) == manifest["semanticCount"] == 49
    assert {icon["file"] for icon in icons} == {path.name for path in SOURCE.glob("*.svg")}
    for icon in icons:
        icon["svg"] = (SOURCE / icon["file"]).read_text(encoding="utf-8")
    # Place the two toggle states next to one another in both previews.
    icons.sort(key=lambda icon: (icon["number"], icon["id"] == "toggle-off"))
    validate(icons)
    OUTPUT.mkdir(parents=True, exist_ok=True)
    contact_sheet(icons)
    html_preview(icons)
    report = dict(svg_assets=len(icons), semantic_icons=49, qt_renders=750,
                  sizes=[16, 20, 24, 32, 48], states=["normal", "disabled", "selected"],
                  checks=["SVG structure", "Qt rendering", "palette", "transparent padding", "no clipping"])
    (OUTPUT / "validation.json").write_text(json.dumps(report, indent=2)+"\n", encoding="utf-8")
    print(json.dumps(report))


if __name__ == "__main__":
    main()
