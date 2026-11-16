"""Infer complete palette items from repeated section bands and pixel gutters.

This is layout evidence, not a detector score or an executable control identity.
No application names, reference screenshots or reviewed coordinates are used.
"""

from __future__ import annotations


def _rect(left, top, right, bottom):
    return dict(left=int(left), top=int(top), right=int(right), bottom=int(bottom),
                width=int(right-left), height=int(bottom-top))


def _area(box):
    return max(1, box['width']*box['height'])


def _overlap(a, b):
    return max(0, min(a['right'], b['right'])-max(a['left'], b['left']))*max(
        0, min(a['bottom'], b['bottom'])-max(a['top'], b['top']))


def _runs(values):
    """Half-open nonzero intervals."""
    import numpy as np

    edges = np.diff(np.r_[False, values, False].astype('int8'))
    return list(zip(np.flatnonzero(edges == 1).tolist(), np.flatnonzero(edges == -1).tolist()))


def discover(image, proposals, texts):
    """Require repeated headers, regular columns/rows and non-text shape support."""
    import cv2
    import numpy as np

    rgb = np.asarray(image.convert('RGB'))
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    neutral = (rgb.max(2).astype('int16')-rgb.min(2)) < 10
    bands = ((gray >= 210) & (gray <= 248) & neutral).astype('uint8')
    bands = cv2.morphologyEx(bands, cv2.MORPH_OPEN, np.ones((1, 100), 'uint8'))
    contours, _ = cv2.findContours(bands, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    groups = []
    for x, y, w, h in sorted((cv2.boundingRect(c) for c in contours), key=lambda b: b[1]):
        if w < 100 or not 12 <= h <= 65:
            continue
        group = next((g for g in groups if abs(g[0][0]-x) <= 3
                      and abs(g[0][2]-w) <= 6 and abs(g[0][3]-h) <= 3), None)
        if group is None:
            groups.append([(x, y, w, h)])
        else:
            group.append((x, y, w, h))
    cells, layouts = [], []
    for headers in groups:
        if len(headers) < 3:
            continue
        x, _, width, header_height = headers[0]
        sections, column_ink, row_steps = [], np.zeros(width, dtype='int64'), []
        for current, following in zip(headers, headers[1:]):
            top, bottom = current[1]+current[3]+1, following[1]-1
            if bottom-top < header_height*.6 or bottom-top > header_height*16:
                continue
            patch = rgb[top:bottom, x+2:x+width-2].astype('int16')
            background = np.median(patch.reshape(-1, 3), axis=0)
            mask = np.max(np.abs(patch-background), axis=2) >= 8
            # Textured/selected backgrounds cannot establish reliable gutters.
            if mask.mean() > .4:
                continue
            # A single thin outline has just its two side pixels on many rows.
            # Keep them connected rather than interpreting top/middle/bottom
            # strokes as three different item rows.
            row_mask = (mask.sum(1) >= 2).astype('uint8')[:, None]
            row_mask = cv2.morphologyEx(row_mask, cv2.MORPH_CLOSE, np.ones((5, 1), 'uint8'))[:, 0]
            rows = _runs(row_mask)
            if not rows:
                continue
            centers = [(a+b)/2 for a, b in rows]
            row_steps.extend(b-a for a, b in zip(centers, centers[1:]))
            column_ink[2:width-2] += mask.sum(0, dtype='int64')
            sections.append(dict(top=top, bottom=bottom, mask=mask, rows=rows))
        columns = _runs(column_ink > 0)
        if len(columns) < 3 or len(sections) < 2 or not row_steps:
            continue
        centers = np.array([(a+b)/2 for a, b in columns])
        widths = np.array([b-a for a, b in columns])
        column_pitch = float(np.median(np.diff(centers)))
        row_pitch = float(np.median(row_steps))
        if not (16 <= row_pitch <= 110 and 20 <= column_pitch <= 150):
            continue
        if (np.max(np.abs(np.diff(centers)-column_pitch)) > column_pitch*.12
                or widths.min() < widths.max()*.5 or widths.max() >= column_pitch*.95):
            continue
        cell_width = int(round(max(widths.max()+2, (widths.max()+column_pitch)/2)))
        max_row_height = max(b-a for s in sections for a, b in s['rows'])
        cell_height = int(round(max(max_row_height+2, row_pitch*.9)))
        if cell_height >= row_pitch or cell_width >= column_pitch:
            continue
        candidate_cells, shape_support, valid_sections = [], 0, []
        for section in sections:
            top, bottom, mask = section['top'], section['bottom'], section['mask']
            count = int(round((bottom-top)/row_pitch))
            if count < 1 or abs(bottom-top-count*row_pitch) > row_pitch*.2:
                continue
            row_centers = [top+(bottom-top-count*row_pitch)/2+(r+.5)*row_pitch for r in range(count)]
            # Every occupied row band must fit one slot; don't force irregular
            # lists or wrapped captions into a component grid.
            if any(not any(c-cell_height/2 <= top+a and top+b <= c+cell_height/2
                           for c in row_centers) for a, b in section['rows']):
                continue
            valid_sections.append(_rect(x, top, x+width, bottom))
            for row, cy in enumerate(row_centers):
                for column, cx in enumerate(centers+x):
                    left, upper = int(round(cx-cell_width/2)), int(round(cy-cell_height/2))
                    box = _rect(left, upper, left+cell_width, upper+cell_height)
                    if (left < x+2 or box['right'] > x+width-2
                            or upper < top or box['bottom'] > bottom):
                        continue  # A clipped slot cannot establish a complete item.
                    local = mask[upper-top:upper-top+cell_height, left-x-2:left-x-2+cell_width]
                    ys, xs = np.nonzero(local)
                    if len(xs) < 6:
                        continue
                    inside_text = [t for t in texts if _overlap(box, t['rect'])/_area(t['rect']) >= .8]
                    residual = local.copy()
                    for text in inside_text:
                        t = text['rect']
                        residual[max(0,t['top']-upper):max(0,t['bottom']-upper),
                                 max(0,t['left']-left):max(0,t['right']-left)] = False
                    ry, rx = np.nonzero(residual)
                    supported = (len(rx) >= 12 and np.ptp(rx) >= column_pitch*.3
                                 and np.ptp(ry) >= row_pitch*.3)
                    shape_support += int(supported)
                    candidate_cells.append(dict(
                        rect=box, pixel_rect=_rect(left+xs.min(), upper+ys.min(),
                                                  left+xs.max()+1, upper+ys.max()+1),
                        row=row, column=column, section_top=top, foreground_pixels=len(xs),
                        nontext_shape_support=bool(supported), ocr_observations=inside_text,
                        supporting_yolo_ids=[i for i,p,_score in proposals
                                             if _overlap(box,p)/_area(p) >= .8]))
        # A regular numeric table alone is insufficient, even when YOLO predicts
        # its words. Require multiple spatially extended non-OCR shapes.
        if len(valid_sections) < 2 or len(candidate_cells) < 6 or shape_support < max(2, len(candidate_cells)*.15):
            continue
        layout_id = len(layouts)+1
        for cell in candidate_cells:
            cell['layout_id'] = layout_id
        cells.extend(candidate_cells)
        layouts.append(dict(id=layout_id, header_rects=[_rect(a,b,a+c,b+d) for a,b,c,d in headers],
                            sections=valid_sections, column_pitch=column_pitch, row_pitch=row_pitch,
                            cell_width=cell_width, cell_height=cell_height,
                            cell_count=len(candidate_cells), nontext_shape_count=shape_support))
    return cells, dict(method='repeated-section-pixel-grid', layouts=layouts, cells=cells)


def partition(proposals, texts, cells, report):
    """Keep intra-cell glyphs as evidence; remove their duplicate operation targets."""
    grouped_ids = {t['raw_id'] for c in cells for t in c['ocr_observations']}
    retained, grouped = [], []
    scopes = [_rect(l['header_rects'][0]['left'], l['header_rects'][0]['top'],
                    l['header_rects'][-1]['right'], l['header_rects'][-1]['bottom'])
              for l in report['layouts']]
    for i, box, score in proposals:
        owners = [c for c in cells if _overlap(box,c['rect'])/min(_area(box),_area(c['rect'])) >= .5]
        if owners and any(_overlap(box,s)/_area(box) >= .9 for s in scopes):
            grouped.append(dict(id=i, rect=box, confidence=score,
                                reason='Palette item contents or enclosing group; represented by complete cells'))
        else:
            retained.append((i,box,score))
    report.update(grouped_yolo=grouped, grouped_ocr_ids=sorted(grouped_ids))
    return retained, [t for t in texts if t['raw_id'] not in grouped_ids]


def elements(cells):
    return [dict(id=0, entity_type='graphic-control', geometry_kind='palette-cell',
                 semantic_name='Palette item candidate', observed_value='', rect=c['rect'],
                 corrected_rect=c['rect'], detected_rect=c['pixel_rect'], geometry_padding=0,
                 source='local-palette-layout', confidence=0, detection_confidence=None,
                 confidence_kind='unscored-local-geometry', yolo_proposal_ids=c['supporting_yolo_ids'],
                 visual_components=[dict(entity_type='text-observation', rect=t['rect'],
                                         observed_value=t['text'], raw_ocr_id=t['raw_id'],
                                         recognition_confidence=t['score']) for t in c['ocr_observations']],
                 geometry_evidence=c, ambiguity='Grid item pixels observed; identity and internal OCR need review',
                 status='observed', uia_verified=False, executable=False, suggested_actions=[]) for c in cells]
