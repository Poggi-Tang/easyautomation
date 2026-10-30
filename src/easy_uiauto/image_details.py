"""Pack populated screenshot regions at native scale for one extra detector pass."""

from __future__ import annotations


def _area(box):
    return max(0, box[2]-box[0])*max(0, box[3]-box[1])


def _overlap(a, b):
    return max(0,min(a[2],b[2])-max(a[0],b[0]))*max(0,min(a[3],b[3])-max(a[1],b[1]))


def plan(image, size=1280):
    """Find content with local edges, never with app names or saved coordinates."""
    import cv2
    import numpy as np

    report = dict(enabled=False, reason='Input already fits the detector at native scale',
                  atlas_size=size, regions=[], omitted_regions=[])
    if max(image.size) <= size:
        return report
    gray=cv2.cvtColor(np.asarray(image.convert('RGB')),cv2.COLOR_RGB2GRAY)
    edges=cv2.Canny(gray,25,70)
    lines=cv2.bitwise_or(cv2.morphologyEx(edges,cv2.MORPH_OPEN,np.ones((1,80),'uint8')),
                        cv2.morphologyEx(edges,cv2.MORPH_OPEN,np.ones((80,1),'uint8')))
    edges[cv2.dilate(lines,np.ones((3,3),'uint8'))>0]=0
    blobs=cv2.dilate(edges,np.ones((13,25),'uint8'))
    contours,_=cv2.findContours(blobs,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    regions=[]
    for contour in contours:
        x,y,w,h=cv2.boundingRect(contour)
        if w*h < 120:
            continue
        regions.append([max(0,x-12),max(0,y-12),min(image.width,x+w+12),min(image.height,y+h+12)])
    # Context margins overlap; merge them before packing so no source region is
    # copied twice and predictions cannot be joined across a patch seam.
    merged=[]
    for box in sorted(regions):
        index=0
        while index < len(merged):
            old=merged[index]
            union=[min(box[0],old[0]),min(box[1],old[1]),max(box[2],old[2]),max(box[3],old[3])]
            if _overlap(box,old)>0 and _area(union)<=(_area(box)+_area(old))*1.5:
                box=union
                merged.pop(index)
                index=0
            else:
                index+=1
        merged.append(box)
    # A bounded detail pass: large dense regions retain the whole-image result.
    # Report every unplaced region instead of pretending this covers the page.
    gutter=16
    x=y=gutter
    shelf=0
    for box in sorted(merged,key=lambda b:(-(b[3]-b[1]),b[1],b[0])):
        w,h=box[2]-box[0],box[3]-box[1]
        if w>size-2*gutter or h>size-2*gutter:
            report['omitted_regions'].append(dict(source_rect=box,reason='Region exceeds native-scale atlas size'))
            continue
        if x+w+gutter>size:
            x=gutter
            y+=shelf+gutter
            shelf=0
        if y+h+gutter>size:
            report['omitted_regions'].append(dict(source_rect=box,reason='One-pass detail budget exhausted'))
            continue
        report['regions'].append(dict(id=len(report['regions'])+1,source_rect=box,atlas_rect=[x,y,x+w,y+h]))
        x+=w+gutter
        shelf=max(shelf,h)
    report.update(enabled=bool(report['regions']),
                  reason='Native-scale populated patches; at most one additional detector pass',
                  source_size=list(image.size),complete_content_coverage=not report['omitted_regions'],
                  parameters=dict(edge_low=25,edge_high=70,context=12,gutter=gutter,max_passes=1))
    return report


def render_atlas(image, report):
    from PIL import Image

    atlas=Image.new('RGB',(report['atlas_size'],report['atlas_size']),'white')
    for region in report['regions']:
        atlas.paste(image.crop(tuple(region['source_rect'])).convert('RGB'),tuple(region['atlas_rect'][:2]))
    return atlas


def merge_predictions(boxes, scores, detail_boxes, detail_scores, report):
    """Append new observations; retain whole-image IDs and audit every patch box."""
    boxes=[list(b) for b in boxes]
    scores=list(scores)
    observations=[]
    added=duplicates=0
    for raw, score in zip(detail_boxes,detail_scores,strict=True):
        area=_area(raw)
        owners=[r for r in report['regions'] if area>0 and _overlap(raw,r['atlas_rect'])/area>=0.95]
        item=dict(atlas_box=list(raw),confidence=float(score))
        if len(owners)!=1:
            observations.append(dict(**item,accepted=False,reason='Prediction crosses patch boundaries or atlas padding'))
            continue
        region=owners[0]
        a,s=region['atlas_rect'],region['source_rect']
        local=[max(raw[0],a[0]),max(raw[1],a[1]),min(raw[2],a[2]),min(raw[3],a[3])]
        box=[local[0]-a[0]+s[0],local[1]-a[1]+s[1],local[2]-a[0]+s[0],local[3]-a[1]+s[1]]
        w,h=box[2]-box[0],box[3]-box[1]
        item.update(region_id=region['id'],source_box=box)
        if not (min(w,h)>=3 and max(w,h)<=96):
            observations.append(dict(**item,accepted=False,reason='Detail pass targets small controls only'))
            continue
        clipped=any((local[i]<=a[i]+2 if i<2 else local[i]>=a[i]-2)
                    and s[i] not in (0,report['source_size'][i%2]) for i in range(4))
        if clipped:
            observations.append(dict(**item,accepted=False,reason='Small control touches an artificial source-crop edge'))
            continue
        matches=[]
        for i, old in enumerate(boxes):
            overlap=_overlap(box,old)
            union=max(1,_area(box)+_area(old)-overlap)
            distance=abs((box[0]+box[2]-old[0]-old[2])/2)+abs((box[1]+box[3]-old[1]-old[3])/2)
            if overlap/union>=0.5 or (overlap/max(1,min(_area(box),_area(old)))>=0.8
                    and max(_area(box),_area(old))<=min(_area(box),_area(old))*2.5
                    and distance<=max(3,min(w,h)*0.35)):
                matches.append(i)
        if matches:
            duplicate=max(matches,key=lambda i:_overlap(box,boxes[i]))
            observations.append(dict(**item,accepted=True,merged_into_yolo_id=duplicate+1))
            duplicates+=1
        else:
            boxes.append(box)
            scores.append(float(score))
            observations.append(dict(**item,accepted=True,added_yolo_id=len(boxes)))
            added+=1
    report.update(observations=observations,added_count=added,duplicate_count=duplicates,
                  rejected_count=sum(not o['accepted'] for o in observations))
    return boxes,scores
