"""Recognize connected model canvases inside UI images from current pixels only."""

from __future__ import annotations


def _area(box):
    return max(1, box['width']*box['height'])


def _overlap(a, b):
    return max(0, min(a['right'], b['right'])-max(a['left'], b['left']))*max(
        0, min(a['bottom'], b['bottom'])-max(a['top'], b['top']))


def _rect(x, y, w, h):
    return dict(left=int(x), top=int(y), right=int(x+w), bottom=int(y+h), width=int(w), height=int(h))


def prepare(image, palette_report):
    """Remove surrounding panel borders from geometry analysis, keeping coordinates."""
    import cv2
    import numpy as np
    from PIL import Image

    rgb = np.asarray(image.convert('RGB'))
    maximum, minimum = rgb.max(2), rgb.min(2)
    light = ((minimum >= 245) & (maximum.astype('int16')-minimum < 30)).astype('uint8')
    light = cv2.erode(light, np.ones((3,3),'uint8'))
    _, backgrounds, stats, _ = cv2.connectedComponentsWithStats(light)
    excluded = [_rect(l['header_rects'][0]['left'], l['header_rects'][0]['top'],
                      l['header_rects'][-1]['right']-l['header_rects'][0]['left'],
                      l['header_rects'][-1]['bottom']-l['header_rects'][0]['top'])
                for l in palette_report.get('layouts', [])]
    canvases = {i: _rect(x,y,w,h) for i,(x,y,w,h,area) in enumerate(stats)
                if i and min(w,h)>=100 and area>=w*h*.65
                and not any(_overlap(_rect(x,y,w,h),p)/(w*h) >= .5 for p in excluded)}
    # An enclosing panel border can otherwise replace all nested component
    # outlines. Analyze only the interior of candidate backgrounds; keep size
    # and coordinates unchanged and still require wire evidence afterwards.
    pixels = np.full_like(rgb,255)
    for b in canvases.values():
        pixels[b['top']:b['bottom'],b['left']:b['right']] = rgb[b['top']:b['bottom'],b['left']:b['right']]
    for b in excluded:
        pixels[b['top']:b['bottom'],b['left']:b['right']] = 255
    return dict(image=Image.fromarray(pixels),backgrounds=backgrounds,stats=stats,
                canvases=canvases,excluded=excluded)


def discover(image, diagram, context):
    """Require separated component bodies joined through exterior wire pixels.

    A light canvas is enabled only when a wire network connects at least three
    bodies. White space, colored buttons, table borders or isolated palette
    icons alone do not enable diagram recovery. Bounds come from background
    components, not names, OCR meanings or reference annotations.
    """
    import cv2
    import numpy as np

    rgb = np.asarray(image.convert('RGB'))
    maximum, minimum = rgb.max(2), rgb.min(2)
    neutral = maximum.astype('int16')-minimum < 30
    backgrounds,stats = context['backgrounds'],context['stats']
    canvases,excluded = context['canvases'],context['excluded']
    graphics = [e for e in diagram['elements'] if e['entity_type']=='graphic-control']
    nodes = []
    for element in graphics:
        box = element['rect']
        if min(box['width'],box['height']) < 8 or max(box['width'],box['height']) < 16:
            continue
        if any(_overlap(box,p)/_area(box) >= .5 for p in excluded):
            continue
        left, top = max(0,box['left']-8), max(0,box['top']-8)
        right, bottom = min(image.width,box['right']+8), min(image.height,box['bottom']+8)
        patch = backgrounds[top:bottom, left:right].copy()
        patch[box['top']-top:box['bottom']-top, box['left']-left:box['right']-left] = 0
        labels, counts = np.unique(patch, return_counts=True)
        choices = [(int(n),int(i)) for i,n in zip(labels,counts) if i in canvases
                   and _overlap(box,canvases[i])/_area(box) >= .9
                   and _area(box) < _area(canvases[i])*.25]
        if choices:
            nodes.append(dict(element_id=element['id'],rect=box,background_id=max(choices)[1]))
    ink = ((maximum < 210) & neutral).astype('uint8')
    # Remove bodies and OCR first; their own borders must not masquerade as
    # connecting wires. Preserve only runs long enough to leave a component.
    for box in [e['rect'] for e in diagram['elements']] + excluded:
        ink[max(0,box['top']-1):min(image.height,box['bottom']+1),
            max(0,box['left']-1):min(image.width,box['right']+1)] = 0
    wires = cv2.bitwise_or(cv2.morphologyEx(ink,cv2.MORPH_OPEN,np.ones((1,9),'uint8')),
                          cv2.morphologyEx(ink,cv2.MORPH_OPEN,np.ones((9,1),'uint8')))
    wires = cv2.morphologyEx(wires,cv2.MORPH_CLOSE,np.ones((3,3),'uint8'))
    _, wire_labels, wire_stats, _ = cv2.connectedComponentsWithStats(wires)
    links = {}
    for node in nodes:
        b = node['rect']
        # Mid-side contacts only: table-cell corner junctions are not ports.
        dx,dy = max(3,round(b['width']*.2)), max(3,round(b['height']*.2))
        bands = [(b['left']-6,b['top']+dy,b['left'],b['bottom']-dy),
                 (b['right'],b['top']+dy,b['right']+6,b['bottom']-dy),
                 (b['left']+dx,b['top']-6,b['right']-dx,b['top']),
                 (b['left']+dx,b['bottom'],b['right']-dx,b['bottom']+6)]
        contacts = set()
        for left,top,right,bottom in bands:
            contacts.update(int(v) for v in np.unique(wire_labels[
                max(0,top):min(image.height,bottom),max(0,left):min(image.width,right)]) if v)
        for wire in contacts:
            links.setdefault((node['background_id'],wire),[]).append(node['element_id'])
    edges, adjacency = [], {}
    by_id = {n['element_id']:n for n in nodes}
    for (background,wire), ids in links.items():
        if len(ids)<2:
            continue
        for pos, first in enumerate(ids):
            a = by_id[first]['rect']
            for second in ids[pos+1:]:
                b = by_id[second]['rect']
                gap = max(0,b['left']-a['right'],a['left']-b['right'],
                          b['top']-a['bottom'],a['top']-b['bottom'])
                if gap<10 or _overlap(a,b):
                    continue
                adjacency.setdefault(first,set()).add(second)
                adjacency.setdefault(second,set()).add(first)
                x,y,w,h,_ = wire_stats[wire]
                edges.append(dict(first=first,second=second,background_id=background,
                                  wire_rect=_rect(x,y,w,h)))
    visited, networks, enabled = set(), [], set()
    for start in adjacency:
        if start in visited:
            continue
        pending, members = [start], set()
        while pending:
            current = pending.pop()
            if current in members:
                continue
            members.add(current)
            pending.extend(adjacency[current]-members)
        visited.update(members)
        if len(members)>=3:
            background = by_id[start]['background_id']
            enabled.add(background)
            networks.append(dict(background_id=background,element_ids=sorted(members)))
    regions = [dict(background_id=i,rect=canvases[i],background_fraction=round(
        float(stats[i,4])/_area(canvases[i]),4)) for i in sorted(enabled)]
    return dict(enabled=bool(regions),method='connected-components-on-light-canvas',
                regions=regions,excluded_regions=excluded,networks=networks,connections=edges,
                evidence_id_space='block-diagram observations before merge',
                geometry_input='Original-size canvas interiors; surrounding panel borders masked white',
                reason='Connected model components on a shared canvas' if regions else
                       'No canvas with three independently bounded components connected by exterior wires')


def in_canvas(box, report):
    return (any(_overlap(box,r['rect'])/_area(box)>=.9 for r in report['regions'])
            and not any(_overlap(box,r)/_area(box)>=.1 for r in report.get('excluded_regions',[])))


def merge(elements, diagram, report):
    """Replace only observations inside confirmed canvases; keep surrounding UI."""
    replaced = [e['id'] for e in elements if in_canvas(e['rect'],report)]
    retained = [dict(e) for e in elements if not in_canvas(e['rect'],report)]
    recovered = [dict(e,diagram_canvas=True) for e in diagram['elements'] if in_canvas(e['rect'],report)]
    # IDs and relationships are local to each fusion result. Remap both sets
    # so text ownership never refers to an unrelated UI control.
    merged = []
    for collection in (retained,recovered):
        remap = {e['id']:len(merged)+i+1 for i,e in enumerate(collection)}
        for e in collection:
            e['id'] = remap[e['id']]
            if 'related_element_id' in e:
                old = e.pop('related_element_id')
                if old in remap:
                    e['related_element_id'] = remap[old]
            if 'relationships' in e:
                e['relationships'] = [dict(r,target_element_id=remap[r['target_element_id']])
                                      for r in e['relationships'] if r.get('target_element_id') in remap]
        merged.extend(collection)
    # Keep the UI export contract: graphics first, then text. Reassign IDs
    # together with references after the two independent observations are joined.
    merged.sort(key=lambda e:(e['entity_type']!='graphic-control',e['rect']['top'],e['rect']['left']))
    final_ids = {e['id']:i+1 for i,e in enumerate(merged)}
    for e in merged:
        e['id'] = final_ids[e['id']]
        if 'related_element_id' in e:
            e['related_element_id'] = final_ids[e['related_element_id']]
        if 'relationships' in e:
            e['relationships'] = [dict(r,target_element_id=final_ids[r['target_element_id']])
                                  for r in e['relationships']]
    report.update(replaced_ui_ids=replaced,
                  recovered_graphic_count=sum(e['entity_type']=='graphic-control' for e in recovered),
                  recovered_text_count=sum(e['entity_type']!='graphic-control' for e in recovered))
    return merged
