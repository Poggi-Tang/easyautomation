"""Local UI evidence for empty-region rejection and repeated-icon recovery."""

from __future__ import annotations

import time


def _rect(x, y, w, h):
    return dict(left=int(x), top=int(y), right=int(x+w), bottom=int(y+h), width=int(w), height=int(h))


def _intersection(a, b):
    return max(0, min(a['right'], b['right'])-max(a['left'], b['left'])) * max(
        0, min(a['bottom'], b['bottom'])-max(a['top'], b['top']))


def _iou(a, b):
    overlap = _intersection(a, b)
    return overlap / max(1, a['width']*a['height']+b['width']*b['height']-overlap)


class UIEvidence:
    def __init__(self, image):
        import cv2
        import numpy as np

        self.rgb = np.asarray(image.convert('RGB'))
        self.gray = cv2.cvtColor(self.rgb, cv2.COLOR_RGB2GRAY)
        self.width, self.height = image.size
        _value, mask = cv2.threshold(self.gray, 225, 255, cv2.THRESH_BINARY_INV)
        contours, _hierarchy = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        frames, rectangles = set(), set()
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            if cv2.contourArea(contour) / max(1, (w-1)*(h-1)) < 0.8:
                continue
            polygon = cv2.approxPolyDP(contour, min(0.035*cv2.arcLength(contour, True), min(w,h)*0.2), True)
            if len(polygon) != 4:
                continue
            if 30 <= w <= 600 and 12 <= h <= 80:
                rectangles.add((x, y, w, h))
            if 8 <= w <= 32 and 8 <= h <= 32 and 0.75 <= w/h <= 1.33:
                frames.add((x, y, w, h))
        self.rectangles = [_rect(*v) for v in sorted(rectangles)]
        # Keep the outer member of a nested outline pair, not both stroke edges.
        self.frames = []
        for value in sorted(frames, key=lambda r: r[2]*r[3], reverse=True):
            box = _rect(*value)
            if any(_intersection(box, old) / (box['width']*box['height']) >= 0.9
                   and old['width'] <= box['width']+4 and old['height'] <= box['height']+4
                   for old in self.frames):
                continue
            self.frames.append(box)

    def pixels(self, box):
        import numpy as np

        x, y = max(0, box['left']-2), max(0, box['top']-2)
        right, bottom = min(self.width, box['right']+2), min(self.height, box['bottom']+2)
        patch = self.rgb[y:bottom, x:right].astype('int16')
        if not patch.size:
            return dict(accepted=False, foreground_pixels=0, reason='Empty image region')
        border = np.concatenate((patch[0], patch[-1], patch[:, 0], patch[:, -1]))
        background = np.median(border, axis=0)
        foreground_mask = np.abs(patch-background).max(2) >= 8
        foreground = int(foreground_mask.sum())
        # The screenshot/window edge alone is not a separate graphic control.
        interior = foreground_mask.copy()
        if x == 0:
            interior[:, :2] = False
        if y == 0:
            interior[:2] = False
        if right == self.width:
            interior[:, -2:] = False
        if bottom == self.height:
            interior[-2:] = False
        edge_only = foreground >= 3 and int(interior.sum()) < 3
        if edge_only:
            return dict(accepted=False, foreground_pixels=foreground,
                        reason='Image boundary pixels alone do not form an independent control')
        return dict(accepted=foreground >= 3, foreground_pixels=foreground,
                    background_rgb=background.tolist(),
                    reason='Visible foreground or border' if foreground >= 3 else
                    'No foreground or border pixels in the proposal and its 2px context')

    def frame_for(self, box):
        if max(box['width'], box['height']) > 42:
            return None
        cx, cy = (box['left']+box['right'])/2, (box['top']+box['bottom'])/2
        choices = [f for f in self.frames
                   if abs((f['left']+f['right'])/2-cx) <= 3
                   and abs((f['top']+f['bottom'])/2-cy) <= 3
                   and (_iou(f, box) >= 0.4 or
                        (_intersection(f, box)/(f['width']*f['height']) >= 0.9
                         and f['width']*f['height'] >= box['width']*box['height']*0.25))]
        return max(choices, key=lambda f: _iou(f, box), default=None)

    def border_for(self, box):
        choices = [f for f in self.rectangles if _iou(f, box) >= 0.35
                   and f['width'] <= box['width']*1.6 and f['height'] <= box['height']*1.6]
        return max(choices, key=lambda f: _iou(f, box), default=None)

    def compact_component(self, box):
        import cv2
        import numpy as np

        if max(box['width'], box['height']) > 40 or min(box['width'], box['height']) < 6:
            return None
        gray = self.gray[box['top']:box['bottom'], box['left']:box['right']]
        _count, _labels, stats, _centers = cv2.connectedComponentsWithStats((gray < 170).astype('uint8'))
        choices = []
        for x, y, w, h, area in stats[1:]:
            if not (3 <= w <= 24 and 3 <= h <= 24 and area >= 5):
                continue
            if x <= 0 or y <= 0 or x+w >= box['width'] or y+h >= box['height']:
                continue
            distance = abs(x+w/2-box['width']/2)+abs(y+h/2-box['height']/2)
            if distance > max(box['width'], box['height'])*0.35:
                continue
            choices.append((distance, _rect(box['left']+x, box['top']+y, w, h)))
        return min(choices, key=lambda p:p[0])[1] if choices else None


def prepare_ui(image, proposals, texts):
    """Use current-image evidence only; no templates from reviewed screenshots."""
    import cv2
    import numpy as np

    start = time.perf_counter()
    evidence = UIEvidence(image)
    from . import ui_layout

    grids, vertical_lines = ui_layout.layout_lines(evidence)
    additions, text_corrections = ui_layout.discover_details(evidence, proposals, texts, vertical_lines)
    symbols = ('>', '<', '›', '‹', 'v', 'V', '×', 'X', '✓', '√')
    components = {}
    for index, box, _score in proposals:
        component = evidence.compact_component(box)
        normal_text = any(_intersection(box, t['rect'])/max(1, box['width']*box['height']) >= 0.5
                          and t['text'].strip() not in symbols for t in texts)
        component_is_text = component and any(
            _intersection(component, t['rect'])/(component['width']*component['height']) >= 0.7
            and t['text'].strip() not in symbols for t in texts)
        components[index] = None if normal_text or component_is_text else component
    known_icons = [r for r in components.values() if r] + evidence.frames
    metadata, rejected, accepted, templates, confirmed_frames = {}, [], [], [], []
    layout_regions = []
    aliases = {}
    for index, box, score in proposals:
        pixels = evidence.pixels(box)
        if not pixels['accepted']:
            rejected.append(dict(id=index, rect=box, confidence=score, **pixels))
            continue
        frame = evidence.frame_for(box)
        region = None if frame else ui_layout.region_kind(box, grids, texts, known_icons)
        if region:
            detail = dict(id=index, rect=box, confidence=score, layout_kind=region,
                          reason='Layout region represented by its independent text and icon contents')
            layout_regions.append(detail)
            rejected.append(detail)
            continue
        duplicate = next(((old_id, old_frame) for old_id, old_frame in confirmed_frames
                          if frame and _iou(frame, old_frame) >= 0.8), None)
        if duplicate:
            primary = duplicate[0]
            aliases.setdefault(primary, []).append(index)
            rejected.append(dict(id=index, rect=box, confidence=score,
                                 reason='Duplicate prediction of one compact control border',
                                 duplicate_of_yolo_id=primary))
            continue
        if frame:
            confirmed_frames.append((index, frame))
        metadata[index] = dict(pixel_evidence=pixels, compact_frame=frame,
                               compact_icon=components[index], closed_border=evidence.border_for(box))
        accepted.append((index, box, score))
    rejected_text = []
    kept_text = []
    for text in texts:
        # A bordered square is independent visual evidence that a lone ASCII
        # OCR glyph such as M was actually the checkbox's artwork.
        literal = text['text'].strip()
        suspicious = literal in ('M', 'm', 'v', 'V', 'X', 'x', '✓', '√', '☑', '□', '■')
        owners = [(index, frame) for index, frame in confirmed_frames
                  if _intersection(text['rect'], frame) / max(1, text['rect']['width']*text['rect']['height']) >= 0.7]
        if suspicious and owners:
            rejected_text.append(dict(id=text['raw_id'], text=text['text'], rect=text['rect'],
                                      reason='OCR glyph inside a visually confirmed compact control border',
                                      supporting_yolo_id=owners[0][0]))
        else:
            kept_text.append(text)
    for index, box, _score in accepted:
        frame = metadata[index]['compact_frame']
        normal_text = any(_intersection(box, t['rect'])/max(1, box['width']*box['height']) >= 0.5
                          and t['text'].strip() not in ('>', '<', 'v', 'V', '×', 'X', '✓', '√')
                          for t in kept_text)
        component = frame if frame else (None if normal_text else components[index])
        if component is None:
            continue
        template = evidence.gray[component['top']:component['bottom'], component['left']:component['right']]
        if template.std() < 8:
            continue
        if any(template.shape == old['pixels'].shape and
               np.abs(template.astype('int16')-old['pixels'].astype('int16')).mean() < 6
               for old in templates):
            continue
        templates.append(dict(source_yolo_id=index, rect=component, pixels=template, boxed=bool(frame)))
        if len(templates) >= 12:
            break
    recovered = []
    covered = [box for _index, box, _score in accepted]
    for seed in templates:
        template = seed['pixels']
        response = cv2.matchTemplate(evidence.gray, template, cv2.TM_CCOEFF_NORMED)
        peaks = (response >= 0.97) & (response >= cv2.dilate(response, np.ones((3,3), dtype='uint8')))
        ys, xs = np.where(peaks)
        positions = sorted(zip(xs.tolist(), ys.tolist()), key=lambda xy: -float(response[xy[1],xy[0]]))[:128]
        for x, y in positions:
            box = _rect(x, y, template.shape[1], template.shape[0])
            if any(_iou(box, old) >= 0.3 or
                   (max(old['width'], old['height']) <= 42
                    and _intersection(box, old)/(box['width']*box['height']) >= 0.75) or
                   (_intersection(box, old)/max(1,min(box['width']*box['height'],old['width']*old['height'])) >= 0.75
                    and max(box['width']*box['height'],old['width']*old['height']) <= box['width']*box['height']*4)
                   for old in covered):
                continue
            if any(_intersection(box, t['rect'])/(box['width']*box['height']) >= 0.7
                   and t['text'].strip() not in ('>', '<', 'v', 'V', '✓', '√', '□')
                   for t in kept_text):
                continue
            if not evidence.pixels(box)['accepted']:
                continue
            match = float(response[y,x])
            recovered.append(dict(rect=box, source_yolo_id=seed['source_yolo_id'],
                                  template_rect=seed['rect'], match_score=round(match,5),
                                  method='same-image-template', boxed=seed['boxed']))
            covered.append(box)
    icon_boxes = [m['compact_frame'] or m['compact_icon'] for m in metadata.values()
                  if m['compact_frame'] or m['compact_icon']]
    icon_boxes.extend(r['rect'] for r in [*recovered, *additions])
    # OCR often reads the exact same arrow twice: once as text, once as a graphic.
    # Only retire a symbolic OCR observation when independent icon evidence agrees.
    final_text = []
    for text in kept_text:
        tb = text['rect']
        matches = [r for r in icon_boxes if _intersection(tb, r)/(r['width']*r['height']) >= 0.8
                   and tb['width']*tb['height'] <= r['width']*r['height']*5]
        if text['text'].strip() in ('>', '<', '›', '‹', '×', '✓', '√') and matches:
            rejected_text.append(dict(id=text['raw_id'], text=text['text'], rect=tb,
                                      reason='OCR symbol duplicates an independently supported graphic icon'))
        else:
            final_text.append(text)
    final_proposals = []
    for index, box, score in accepted:
        meta = metadata[index]
        list_row = any(r['layout_kind'] == 'repeated-list-region'
                       and box['height'] <= 50 and box['width'] >= r['rect']['width']*0.6
                       and _intersection(box, r['rect'])/(box['width']*box['height']) >= 0.85
                       for r in layout_regions)
        if list_row:
            detail = dict(id=index, rect=box, confidence=score, layout_kind='list-row',
                          reason='List row represented by its independent text and icon contents')
            layout_regions.append(detail)
            rejected.append(detail)
        elif not (meta['compact_frame'] or meta['compact_icon'] or meta['closed_border']) and ui_layout.content_only(
                evidence, box, final_text, icon_boxes):
            rejected.append(dict(id=index, rect=box, confidence=score,
                                 reason='OCR and independent icons explain all foreground in this region'))
        else:
            final_proposals.append((index, box, score))
    report = dict(enabled=True, rejected_yolo=rejected,
                  protected_yolo_ids=[i for i,m in metadata.items() if m['compact_frame']],
                  compact_frame_count=len(evidence.frames),
                  merged_yolo_ids=aliases, rejected_ocr=rejected_text,
                  grids=grids, layout_regions=layout_regions,
                  geometry_additions=additions, text_corrections=text_corrections,
                  template_count=len(templates), recovered=recovered, added_count=len(recovered),
                  parameters=dict(context=2, minimum_foreground_pixels=3, pixel_contrast=8,
                                  template_threshold=0.97, max_templates=12, max_matches_per_template=128),
                  elapsed_seconds=round(time.perf_counter()-start,4))
    return final_proposals, final_text, metadata, recovered, report
