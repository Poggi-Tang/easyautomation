"""Default UI annotation must recover connected diagrams without converting forms."""

import copy
import json
from pathlib import Path

from PIL import Image, ImageDraw
import pytest

from easy_uiauto import image_detection, ui_diagram
from easy_uiauto.mcp import rectangle_snap


def _fixture(offset=(0,0), scale=1, connected=True, count=3):
    image=Image.new('RGB',(1100,760),'#eeeeee')
    draw=ImageDraw.Draw(image)
    dx,dy=offset
    draw.rectangle((250+dx,100+dy,750+dx,500+dy),fill='white',outline='black')
    boxes,ocr_boxes,words,bodies=[],[],[],[]
    for i in range(count):
        x,y=280+160*i+dx,180+dy
        draw.rectangle((x,y,x+80,y+60),fill='white',outline='black')
        draw.text((x+20,y+20),'Block',fill='black')
        bodies.append((x,y,x+81,y+61))
        ocr_boxes.append((x+20,y+20,x+55,y+35));words.append('Block')
        boxes.append((x,y+53,x+81,y+65))  # incomplete bottom-edge detection
        if connected and i<count-1:
            draw.line((x+80,y+30,x+160,y+30),fill='black')
    # UI controls outside the canvas must keep their original fusion behavior.
    draw.rectangle((20,25,130,55),fill='white',outline='black')
    draw.text((45,32),'Run',fill='black')
    boxes.append((19,24,132,57));ocr_boxes.append((45,32,66,45));words.append('Run')
    scaled=lambda p:tuple(round(v*scale) for v in p)
    if scale!=1:
        image=image.resize((round(image.width*scale),round(image.height*scale)),Image.Resampling.NEAREST)
    return image,dict(boxes=[scaled(b) for b in boxes],scores=[.8]*len(boxes)),dict(
        boxes=[[[a,b],[c,b],[c,d],[a,d]] for a,b,c,d in map(scaled,ocr_boxes)],
        texts=words,scores=[.99]*len(words)),[scaled(b) for b in bodies]


@pytest.mark.parametrize('offset,scale',[((0,0),1),((117,93),1),((21,12),1.5)])
def test_embedded_diagram_recovers_complete_bodies_and_preserves_other_ui(offset,scale):
    image,yolo,ocr,bodies=_fixture(offset,scale)
    raw=copy.deepcopy((yolo,ocr))
    result=image_detection.combine_detections(image,yolo,ocr,profile='ui')
    assert (yolo,ocr)==raw
    report=result['ui_fusion']['diagram_recovery']
    assert report['enabled'] and report['recovered_graphic_count']==3
    recovered=[e for e in result['elements'] if e.get('diagram_canvas') and e['entity_type']=='graphic-control']
    for left,top,right,bottom in bodies:
        matches=[e for e in recovered if e['rect']['left']<=left and e['rect']['top']<=top
                 and e['rect']['right']>=right and e['rect']['bottom']>=bottom]
        assert len(matches)==1
        assert matches[0]['rect']['width']<90*scale  # wire is not part of the box
    outside=[e for e in result['elements'] if e['rect']['top']<60*scale]
    assert any(e['entity_type']=='graphic-control' for e in outside)
    assert any(e['observed_value']=='Run' for e in outside)
    assert all(not e.get('diagram_canvas') for e in outside)
    ids={e['id'] for e in result['elements']}
    assert len(ids)==len(result['elements'])
    for e in result['elements']:
        if 'related_element_id' in e:
            assert e['related_element_id'] in ids
            owner=next(g for g in result['elements'] if g['id']==e['related_element_id'])
            assert owner['entity_type']=='graphic-control'
            assert ui_diagram._overlap(e['rect'],owner['rect'])>0


@pytest.mark.parametrize('connected,count',[(False,3),(True,2)])
def test_unconnected_cards_or_a_single_connector_do_not_enable_diagram_recovery(connected,count):
    image,yolo,ocr,_=_fixture(connected=connected,count=count)
    result=image_detection.combine_detections(image,yolo,ocr,profile='ui')
    assert not result['ui_fusion']['diagram_recovery']['enabled']


def test_form_table_borders_do_not_count_as_component_connections():
    image=Image.new('RGB',(650,420),'white');draw=ImageDraw.Draw(image)
    yolo=dict(boxes=[],scores=[]);ocr=dict(boxes=[],texts=[],scores=[])
    for row in range(4):
        for column in range(4):
            x,y=40+column*120,40+row*65
            draw.rectangle((x,y,x+120,y+65),outline='black')
            draw.text((x+25,y+25),'Value',fill='black')
            yolo['boxes'].append([x,y,x+120,y+65]);yolo['scores'].append(.8)
            ocr['boxes'].append([[x+25,y+25],[x+55,y+25],[x+55,y+38],[x+25,y+38]])
            ocr['texts'].append('Value');ocr['scores'].append(.99)
    r=image_detection.combine_detections(image,yolo,ocr,profile='ui')
    assert not r['ui_fusion']['diagram_recovery']['enabled']


def test_palette_and_model_canvas_coexist_without_merging_their_controls():
    from test_ui_palette import _fixture as palette_fixture

    image,yolo,ocr,bodies=_fixture()
    larger=Image.new('RGB',(1500,760),'#eeeeee');larger.paste(image,(0,0))
    palette,centers,boxes,other_ocr=palette_fixture()
    larger.paste(palette,(1050,120))
    yolo['boxes'].extend([[a+1050,b+120,c+1050,d+120] for a,b,c,d in boxes])
    yolo['scores'].extend([.8]*len(boxes))
    ocr['boxes'].extend([[[x+1050,y+120] for x,y in quad] for quad in other_ocr['boxes']])
    ocr['texts'].extend(other_ocr['texts']);ocr['scores'].extend(other_ocr['scores'])
    r=image_detection.combine_detections(larger,yolo,ocr,profile='ui')
    assert r['ui_fusion']['diagram_recovery']['recovered_graphic_count']==3
    cells=[e for e in r['elements'] if e.get('geometry_kind')=='palette-cell']
    assert len(cells)==len(centers)==14
    assert all(not e.get('diagram_canvas') for e in cells)
    assert len({(c['rect']['width'],c['rect']['height']) for c in cells})==1


def test_ieee9_default_ui_recovers_the_published_complete_component_bounds():
    # Public original + raw detector evidence are inputs. Published annotations
    # are used only afterwards to assess regression, never by the pipeline.
    base=Path(__file__).resolve().parents[1]/'docs/examples/IEEE9'
    yolo=json.loads((base/'results/raw-yolo.json').read_text(encoding='utf8'))
    ocr=json.loads((base/'results/raw-ocr.json').read_text(encoding='utf8'))
    reference=json.loads((base/'results/elements.json').read_text(encoding='utf8'))
    image=Image.open(base/'IEEE9.bmp')
    result=image_detection.combine_detections(image,yolo,ocr,profile='ui')
    elements,_=rectangle_snap.snap_rectangles(image,result['elements'])
    graphics=[e for e in elements if e['entity_type']=='graphic-control']
    expected=[e for e in reference['elements'] if e['entity_type']=='graphic-control']
    assert len(graphics)==len(expected)==42
    assert [e['rect'] for e in graphics]==[e['rect'] for e in expected]
    assert result['text_count']==102
    assert result['ui_fusion']['diagram_recovery']['enabled']
    assert all(not e['executable'] for e in graphics)
