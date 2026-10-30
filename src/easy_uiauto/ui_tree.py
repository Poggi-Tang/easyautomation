"""Recover repeated tree prefixes without treating complete OCR rows as icons."""

from __future__ import annotations

from .ui_image_fusion import _intersection, _iou, _rect


def discover(evidence, proposals, texts, known):
    import cv2
    import numpy as np

    mask=(evidence.gray<200).astype('uint8')
    count,labels,stats,_centers=cv2.connectedComponentsWithStats(mask)
    candidates=[]
    for component in range(1,count):
        x,y,w,h,area=map(int,stats[component])
        if not (3<=min(w,h) and max(w,h)<=18 and area<=w*h*0.75):
            continue
        pixels=labels[y:y+h,x:x+w]==component
        direction=None
        for axis,name in ((1,'right'),(0,'down')):
            rows=pixels if axis==1 else pixels.T
            profile=np.array([np.where(row)[0].mean() if row.any() else 0 for row in rows])
            if len(profile)<5:
                continue
            end=max(1,len(profile)//4)
            peak=profile[len(profile)//2-1:len(profile)//2+2].mean()
            edge=(profile[:end].mean()+profile[-end:].mean())/2
            if peak-edge>=(rows.shape[1]-1)*0.5 and np.abs(profile-profile[::-1]).mean()<=1.5:
                direction=name
                break
        if direction is None:
            continue
        box=_rect(x,y,w,h)
        owners=[t for t in texts if len(t['text'].strip())>=2
                and any(c.isalnum() for c in t['text'])
                and abs((t['rect']['top']+t['rect']['bottom'])/2-(y+h/2))<=t['rect']['height']*0.6
                and x<=t['rect']['left']+t['rect']['height']*0.6
                and x+w>=t['rect']['left']-t['rect']['height']*2
                and t['rect']['right']>=x+w+t['rect']['height']*0.8]
        support=[i for i,p,_score in proposals if p['height']<=48
                 and _intersection(box,p)/(w*h)>=0.6]
        if owners:
            owner=min(owners,key=lambda t:abs(t['rect']['left']-x))
            after=mask[owner['rect']['top']:owner['rect']['bottom'],x+w:owner['rect']['right']]
            following=np.where(after.any(0))[0]
            if len(following) and int(following[0])>=3:
                candidates.append(dict(rect=box,direction=direction,owner=owner,support=support))
    repeated=[]
    for c in candidates:
        a=c['rect']
        if any(other is not c and abs(other['rect']['left']-a['left'])<=max(3,a['height']*4)
               and a['height']<abs(other['rect']['top']-a['top'])<=max(a['height'],other['rect']['height'])*25
               for other in candidates):
            repeated.append(c)
    additions,corrections,columns=[],[],[]
    occupied=list(known)

    def add(box,kind,support):
        if any(_iou(box,old)>=0.45 or (_intersection(box,old)/max(1,box['width']*box['height'])>=0.8
               and old['width']*old['height']<=box['width']*box['height']*3) for old in occupied):
            return
        additions.append(dict(rect=box,geometry_kind=kind,supporting_yolo_ids=support,
                              method='repeated-tree-prefix-pixels'))
        occupied.append(box)

    def trim(text,right,strip_symbol=False):
        if text['rect']['left']>=right:
            return
        before=dict(rect=dict(text['rect']),text=text['text'],quad=text['quad'])
        row=mask[text['rect']['top']:text['rect']['bottom'],right:text['rect']['right']]
        ink=np.where(row.any(0))[0]
        if not len(ink):
            return
        left=right+int(ink[0])
        if strip_symbol and text['text'].lstrip().startswith(('>','√','⌄','v','V')):
            text['text']=text['text'].lstrip()[1:].lstrip()
        b=text['rect']
        text['rect']=_rect(left,b['top'],b['right']-left,b['height'])
        text['quad']=[[text['rect'][a],text['rect'][b]] for a,b in
                      (('left','top'),('right','top'),('right','bottom'),('left','bottom'))]
        record=dict(method='separate-tree-prefix',before=before,
                    previous=text.get('layout_correction'),prefix_right=right)
        text['layout_correction']=record
        corrections.append(dict(raw_ocr_id=text['raw_id'],**record))

    owners_with_prefix=set()
    for c in repeated:
        box,text=c['rect'],c['owner']
        add(box,'tree-chevron',c['support'])
        y1,y2=text['rect']['top'],text['rect']['bottom']
        start,end=box['right']+1,min(text['rect']['right'],box['right']+text['rect']['height']*3)
        row=mask[y1:y2,start:end]
        runs=np.where(np.diff(np.pad(row.any(0).astype('int8'),(1,1))))[0]
        prefix=None
        if len(runs)>=4:
            a,b=int(runs[0]),int(runs[1])
            yy,xx=np.where(row[:,a:b])
            if len(xx) and a>=3 and 5<=b-a<=text['rect']['height']*1.3 and int(runs[2])-b>=3:
                prefix=_rect(start+a,y1+int(yy.min()),b-a,int(yy.max()-yy.min()+1))
                if 0.6<=prefix['width']/prefix['height']<=1.6:
                    add(prefix,'tree-prefix-icon',c['support'])
                    columns.append((prefix['left'],prefix['right']))
                    owners_with_prefix.add(id(text))
                else:
                    prefix=None
        trim(text,(prefix or box)['right']+1,strip_symbol=True)
    owned={id(c['owner']):c for c in repeated}
    # Leaf rows may have the same icon column without an expansion arrow.
    for text in texts:
        if id(text) in owners_with_prefix:
            continue
        b=text['rect']
        for left,right in sorted(set(columns)):
            if not (b['left']<=left+2 and b['right']>right+b['height']):
                continue
            if id(text) in owned:
                arrow=owned[id(text)]['rect']
                before=mask[b['top']:b['bottom'],arrow['right']+1:max(arrow['right']+1,left-2)]
                if left-arrow['right']<3 or left-arrow['right']>b['height']*2 or before.any():
                    continue
            patch=mask[b['top']:b['bottom'],max(0,left-2):right+2]
            yy,xx=np.where(patch)
            if not len(xx):
                continue
            box=_rect(max(0,left-2)+int(xx.min()),b['top']+int(yy.min()),
                      int(xx.max()-xx.min()+1),int(yy.max()-yy.min()+1))
            next_ink=np.where(mask[b['top']:b['bottom'],box['right']:b['right']].any(0))[0]
            support=[i for i,p,_s in proposals if p['height']<=48 and
                     _intersection(box,p)/(box['width']*box['height'])>=0.6]
            if (support and 5<=min(box['width'],box['height']) and 0.6<=box['width']/box['height']<=1.6
                    and len(next_ink) and int(next_ink[0])>=3):
                add(box,'tree-prefix-icon',support)
                trim(text,box['right']+1)
                break
    return additions,corrections,dict(chevron_candidates=len(candidates),repeated_chevrons=len(repeated),
                                      icon_columns=[list(c) for c in sorted(set(columns))])
