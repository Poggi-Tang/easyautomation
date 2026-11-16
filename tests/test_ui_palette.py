"""Palette cells are inferred from pixels, including missed and OCR-confused icons."""

from PIL import Image, ImageDraw
import pytest

from easy_uiauto import image_detection
from easy_uiauto.mcp import image_analysis, rectangle_snap


def _fixture(offset=(37,29), scale=1, text_only=False, collapsed=False):
    image=Image.new('RGB',(430,480),'white')
    draw=ImageDraw.Draw(image)
    x,y=offset
    centers,boxes,ocr_boxes,ocr_texts=[],[],[],[]
    for section,counts in enumerate(((4,3),(4,2),(1,))):
        draw.rectangle((x,y,x+260,y+23),fill=(235,235,235))
        y+=25
        for row,count in enumerate(counts if not collapsed else ()):
            for col in range(count):
                cx,cy=x+30+col*52,y+19+row*38
                centers.append((cx,cy))
                if text_only:
                    draw.text((cx-12,cy-5),'1000',fill='black')
                    ocr_boxes.append((cx-12,cy-5,cx+12,cy+6))
                    ocr_texts.append('1000')
                else:
                    # Complete item includes separated terminals and an inner
                    # rectangular symbol that the generic snapper must not use.
                    draw.line((cx-19,cy,cx+19,cy),fill='black')
                    draw.rectangle((cx-15,cy-12,cx+15,cy+12),outline='black')
                    if col==0:
                        boxes.extend([(cx-15,cy-12,cx+16,cy+13),(cx-3,cy-3,cx+3,cy+3)])
                    if section==1 and row==0 and col==2:
                        ocr_boxes.append((cx-19,cy-12,cx+20,cy+13))
                        ocr_texts.append('王')
            if not text_only and not collapsed and count>1:
                boxes.append((x,y+row*38,x+260,y+(row+1)*38)) # enclosing group
        if not collapsed:
            y+=len(counts)*38+1
    draw.rectangle((x,y,x+260,y+23),fill=(235,235,235))
    if scale!=1:
        image=image.resize((round(image.width*scale),round(image.height*scale)),Image.Resampling.NEAREST)
    scaled=lambda b:tuple(round(v*scale) for v in b)
    return image,[scaled(p) for p in centers],[scaled(b) for b in boxes],dict(
        boxes=[[[a,b],[c,b],[c,d],[a,d]] for a,b,c,d in map(scaled,ocr_boxes)],
        texts=ocr_texts,scores=[.99]*len(ocr_texts))


@pytest.mark.parametrize('offset,scale',[((37,29),1),((89,71),1),((37,29),1.5)])
def test_missing_cells_recovered_once_with_uniform_complete_geometry(offset,scale):
    image,centers,boxes,ocr=_fixture(offset,scale)
    result=image_detection.combine_detections(image,dict(boxes=boxes,scores=[.8]*len(boxes)),ocr,profile='ui')
    cells=[e for e in result['elements'] if e.get('geometry_kind')=='palette-cell']
    assert len(cells)==len(centers)==14  # incomplete rows do not invent empty cells
    assert len({(c['rect']['width'],c['rect']['height']) for c in cells})==1
    for x,y in centers:
        matches=[c for c in cells if c['rect']['left']<x<c['rect']['right'] and c['rect']['top']<y<c['rect']['bottom']]
        assert len(matches)==1
        assert matches[0]['rect']['width']>=39*scale
        assert matches[0]['rect']['height']>=25*scale
        # No duplicate fragment, whole-group box or OCR operation target.
        assert sum(e['rect']['left']<x<e['rect']['right'] and e['rect']['top']<y<e['rect']['bottom']
                   for e in result['elements'])==1
    attached=[v for c in cells for v in c['visual_components']]
    assert [v['observed_value'] for v in attached]==['王']
    assert all(not c['executable'] and c['detection_confidence'] is None for c in cells)
    # Re-completing and snapping must not shrink cells to their inner symbols.
    completed=image_analysis.complete_element_bounds(cells,*image.size,padding=8)
    snapped,_report=rectangle_snap.snap_rectangles(image,completed,box_padding=8)
    assert [c['rect'] for c in snapped]==[c['rect'] for c in cells]


@pytest.mark.parametrize('text_only,collapsed',[(True,False),(False,True)])
def test_regular_text_tables_and_collapsed_sections_are_not_palette_items(text_only,collapsed):
    image,_centers,boxes,ocr=_fixture(text_only=text_only,collapsed=collapsed)
    result=image_detection.combine_detections(image,dict(boxes=boxes,scores=[.8]*len(boxes)),ocr,profile='ui')
    assert not result['ui_fusion']['palette_layout']['cells']
