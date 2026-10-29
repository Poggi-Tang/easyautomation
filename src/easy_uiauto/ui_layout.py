"""Pixel evidence for UI layout regions and text-adjacent graphic details.

Layout regions remain in the audit report, rather than acquiring a second
graphic ID on top of the text and icons that already describe their contents.
"""

from __future__ import annotations

from .ui_image_fusion import _intersection, _rect


def layout_lines(evidence):
    import cv2
    import numpy as np

    mask = (evidence.gray < 225).astype('uint8') * 255
    horizontal = cv2.morphologyEx(mask, cv2.MORPH_OPEN,
                                  np.ones((1, max(40, evidence.width // 12)), 'uint8'))
    vertical = cv2.morphologyEx(mask, cv2.MORPH_OPEN,
                                np.ones((max(60, evidence.height // 8), 1), 'uint8'))

    def bounds(image):
        contours, _ = cv2.findContours(image, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        return [_rect(*cv2.boundingRect(c)) for c in contours]

    hs, vs = bounds(horizontal), bounds(vertical)
    groups = []
    for line in hs:
        if line['height'] > 3 or line['width'] < 100:
            continue
        group = next((g for g in groups if abs(g[0]['left']-line['left']) <= 3
                      and abs(g[0]['right']-line['right']) <= 3), None)
        if group is None:
            groups.append([line])
        else:
            group.append(line)
    grids = []
    for lines in groups:
        ys = sorted({r['top'] for r in lines})
        if len(ys) < 4:
            continue
        left, right = lines[0]['left'], lines[0]['right']
        xs = sorted({r['left'] for r in vs if r['width'] <= 3
                     and left-3 <= r['left'] <= right+3
                     and r['top'] <= ys[0]+4 and r['bottom'] >= ys[0]+(ys[-1]-ys[0])*0.85})
        if len(xs) >= 3:
            grids.append(dict(rect=_rect(min(xs), ys[0], max(xs)-min(xs)+1, ys[-1]-ys[0]+1),
                              vertical_lines=xs, horizontal_lines=ys))
    return grids, vs


def region_kind(box, grids, texts, icons):
    """Only classify grid-aligned fragments or content-only row/container boxes."""
    for grid in grids:
        if _intersection(box, grid['rect']) / (box['width']*box['height']) < 0.95:
            continue
        xs, ys = grid['vertical_lines'], grid['horizontal_lines']
        if (min(abs(box['top']-v) for v in ys) <= 4
                and min(abs(box['bottom']-v) for v in ys) <= 4):
            cell = (min(abs(box['left']-v) for v in xs) <= 4
                    and min(abs(box['right']-v) for v in xs) <= 4)
            return 'grid-cell' if cell else 'grid-row-fragment'
    contained = [t for t in texts if _intersection(box, t['rect']) /
                 (t['rect']['width']*t['rect']['height']) >= 0.8]
    rows = []
    for text in sorted(contained, key=lambda t: t['rect']['top']):
        cy = (text['rect']['top']+text['rect']['bottom']) / 2
        if not rows or cy-rows[-1] > text['rect']['height']:
            rows.append(cy)
    icon_count = sum(_intersection(box, r) / (r['width']*r['height']) >= 0.8 for r in icons)
    if len(rows) >= 4 and icon_count >= 3:
        return 'repeated-list-region'
    return None


def content_only(evidence, box, texts, icons):
    """Do not emit another row box when OCR and icons explain all dark pixels."""
    patch = evidence.gray[box['top']:box['bottom'], box['left']:box['right']] < 170
    related = [t for t in texts if _intersection(box, t['rect']) /
               (t['rect']['width']*t['rect']['height']) >= 0.7]
    if not related:
        return False
    for rect in [*(t['rect'] for t in texts), *icons]:
        left, top = max(0, rect['left']-box['left']-1), max(0, rect['top']-box['top']-1)
        right, bottom = min(box['width'], rect['right']-box['left']+1), min(
            box['height'], rect['bottom']-box['top']+1)
        if right > left and bottom > top:
            patch[top:bottom, left:right] = False
    return int(patch.sum()) <= 2


def discover_details(evidence, proposals, texts, vertical_lines):
    """Recover a supported scroll thumb or a magnifier merged into a CJK label."""
    import cv2
    import numpy as np

    additions, corrections = [], []
    for bar in vertical_lines:
        if not (4 <= bar['width'] <= 16 and bar['height'] >= max(60, bar['width']*6)):
            continue
        patch = evidence.gray[bar['top']:bar['bottom'], bar['left']:bar['right']]
        # A filled, light gray thumb has width; a table divider has only a stroke.
        if not (150 <= float(np.median(patch)) <= 235 and float(patch.std()) < 30):
            continue
        support = [i for i, box, _ in proposals
                   if abs(box['right']-bar['left']) <= 16
                   and _intersection(_rect(box['left'], box['top'], box['width']+16, box['height']), bar)
                   / (bar['width']*bar['height']) >= 0.8 and box['width'] > 60]
        if support:
            additions.append(dict(rect=bar, geometry_kind='scrollbar-thumb',
                                  supporting_yolo_ids=support, method='filled-scroll-thumb'))
    count, labels, stats, _ = cv2.connectedComponentsWithStats((evidence.gray < 170).astype('uint8'))
    for component_id in range(1, count):
        x, y, w, h, _area = map(int, stats[component_id])
        if not (10 <= w <= 28 and 10 <= h <= 28 and 0.8 <= w/h <= 1.2):
            continue
        box = _rect(x, y, w, h)
        owners = [t for t in texts if t['text'].strip().endswith(('Q', 'O'))
                  and any('\u4e00' <= c <= '\u9fff' for c in t['text'])
                  and _intersection(box, t['rect'])/(w*h) >= 0.9
                  and abs(t['rect']['right']-box['right']) <= 3
                  and t['rect']['width'] >= w*3]
        if not owners:
            continue
        mask = (labels[y:y+h, x:x+w] == component_id).astype('uint8') * 255
        contours, hierarchy = cv2.findContours(mask, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
        holes = [c for i, c in enumerate(contours) if hierarchy[0][i][3] >= 0]
        hole = max(holes, key=cv2.contourArea, default=None)
        if hole is None:
            continue
        hx, hy, hw, hh = cv2.boundingRect(hole)
        # The ring sits above/left of the tail. A plain O/0 has a centered hole.
        if not (hw >= w*0.4 and hh >= h*0.4 and 0.7 <= hw/hh <= 1.4
                and hx+hw/2 < w*0.49 and hy+hh/2 < h*0.49
                and mask[-2:, -2:].any()):
            continue
        support = [i for i, p, _ in proposals if _intersection(box, p)/(w*h) >= 0.9]
        if not support:
            continue
        owner = owners[0]
        before = dict(rect=dict(owner['rect']), text=owner['text'])
        owner['rect'] = _rect(owner['rect']['left'], owner['rect']['top'],
                              x-owner['rect']['left']-1, owner['rect']['height'])
        owner['text'] = owner['text'].rstrip()[:-1].rstrip()
        owner['quad'] = [[owner['rect'][a], owner['rect'][b]]
                         for a, b in (('left','top'), ('right','top'), ('right','bottom'), ('left','bottom'))]
        owner['layout_correction'] = dict(method='separate-magnifier-glyph', before=before,
                                          graphic_rect=box)
        corrections.append(dict(raw_ocr_id=owner['raw_id'], **owner['layout_correction']))
        additions.append(dict(rect=box, geometry_kind='magnifier-icon',
                              supporting_yolo_ids=support, method='separate-magnifier-glyph'))
    return additions, corrections
