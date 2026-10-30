"""UI fusion must retain real controls, reject blank predictions and audit recovery."""

from PIL import Image, ImageDraw
import pytest

from easy_uiauto import image_detection
from easy_uiauto.ui_image_fusion import UIEvidence, _rect


def _ocr(boxes=(), texts=()):
    return dict(boxes=[[[a,b],[c,b],[c,d],[a,d]] for a,b,c,d in boxes],
                texts=list(texts), scores=[0.99]*len(texts))


def _combine(image, boxes=(), texts=None):
    return image_detection.combine_detections(image, dict(boxes=list(boxes), scores=[0.8]*len(boxes)),
                                               texts or _ocr(), profile='ui')


def test_pure_blank_is_rejected_but_empty_bordered_input_is_retained():
    image = Image.new('RGB',(260,180),'white')
    ImageDraw.Draw(image).rectangle((20,20,220,50),outline=(220,220,220))
    result = _combine(image, [(18,18,223,53),(20,90,220,165)])
    assert result['graphic_count']==1 and result['blank_rejected_graphics']==1
    assert result['elements'][0]['yolo_proposal_ids']==[1]
    assert 'No foreground or border' in result['ui_fusion']['rejected_yolo'][0]['reason']


@pytest.mark.parametrize('line',[(68,50,68,170),(171,50,171,170),
                                  (40,78,200,78),(40,141,200,141)])
def test_a_nearby_panel_line_cannot_validate_a_blank_prediction(line):
    image=Image.new('RGB',(260,210),'white')
    ImageDraw.Draw(image).line(line,fill=(191,191,191))
    result=_combine(image,[(70,80,170,140)])
    assert result['graphic_count']==0
    rejected=result['ui_fusion']['rejected_yolo'][0]
    assert rejected['inside_foreground_pixels']==0
    assert rejected['context_foreground_pixels']>0
    assert rejected['pixel_support']=='none'


@pytest.mark.parametrize('outer,inner',[
    ((20,20,220,50),(21,21,220,50)),
    ((20,20,35,35),(21,21,35,35)),
    ((20,20,900,50),(21,21,900,50)),
])
@pytest.mark.parametrize('gray',[220,245])
def test_empty_inputs_and_checkboxes_keep_their_matching_closed_border(outer,inner,gray):
    image=Image.new('RGB',(max(260,outer[2]+20),180),'white')
    ImageDraw.Draw(image).rectangle(outer,outline=(gray,gray,gray))
    evidence=UIEvidence(image).pixels(_rect(inner[0],inner[1],inner[2]-inner[0],inner[3]-inner[1]))
    assert evidence['accepted'] and evidence['inside_foreground_pixels']==0
    assert evidence['pixel_support']=='closed-border-context'
    assert evidence['context_closed_border']
    result=_combine(image,[inner])
    assert result['graphic_count']==1


def test_three_external_sides_do_not_count_as_a_closed_input():
    image=Image.new('RGB',(260,180),'white')
    ImageDraw.Draw(image).line([(20,50),(20,20),(220,20),(220,50)],fill=(220,220,220))
    evidence=UIEvidence(image).pixels(_rect(21,21,199,29))
    assert not evidence['accepted'] and evidence['inside_foreground_pixels']==0
    assert evidence['context_foreground_pixels']>0 and evidence['context_closed_border'] is None


def test_checkbox_is_not_deleted_when_ocr_misreads_it_as_m():
    image = Image.new('RGB',(160,90),'white')
    draw = ImageDraw.Draw(image)
    draw.rectangle((20,20,35,35),outline='black')
    draw.line([(23,27),(27,31),(32,23)],fill='blue',width=2)
    draw.text((60,20),'Value',fill='black')
    ocr = _ocr([(20,20,36,36),(60,20,90,33)],['M','Value'])
    result = _combine(image,[(18,18,38,38)],ocr)
    assert result['graphic_count']==1 and result['text_count']==1
    assert result['elements'][0]['geometry_kind']=='compact-frame'
    assert result['elements'][1]['observed_value']=='Value'
    assert result['ui_fusion']['rejected_ocr'][0]['text']=='M'
    assert ocr['texts']==['M','Value']


def test_identical_icons_are_recovered_from_current_image_without_duplicating_seed():
    image = Image.new('RGB',(180,80),'white')
    draw = ImageDraw.Draw(image)
    for x in (20,70,120):
        draw.rectangle((x,20,x+15,35),outline='black')
    result = _combine(image,[(18,18,38,38)])
    assert result['graphic_count']==3 and result['template_recovered_graphics']==2
    recovered = [e for e in result['elements'] if e['source']=='local-template-recovery']
    assert len(recovered)==2
    assert all(e['confidence_kind']=='template-correlation' and not e['yolo_proposal_ids'] for e in recovered)
    assert all(e['template_evidence']['source_yolo_id']==1 for e in recovered)


def test_recovery_requires_a_detected_seed_and_does_not_invent_controls_from_frames():
    image = Image.new('RGB',(180,80),'white')
    for x in (20,70,120):
        ImageDraw.Draw(image).rectangle((x,20,x+15,35),outline='black')
    result = _combine(image)
    assert result['graphic_count']==result['template_recovered_graphics']==0


def test_same_frame_predictions_are_merged_with_all_raw_yolo_ids_retained():
    image = Image.new('RGB',(100,80),'white')
    ImageDraw.Draw(image).rectangle((20,20,35,35),outline='black')
    result = _combine(image,[(18,18,38,38),(16,16,40,40)])
    assert result['graphic_count']==1
    assert result['elements'][0]['yolo_proposal_ids']==[1,2]
    assert 'Duplicate' in result['ui_fusion']['rejected_yolo'][0]['reason']


def test_regular_text_is_kept_as_text_without_graphic_duplication():
    image = Image.new('RGB',(180,80),'white')
    ImageDraw.Draw(image).text((20,20),'Label',fill='black')
    result = _combine(image,[(19,19,56,34)],_ocr([(19,19,56,34)],['Label']))
    assert result['graphic_count']==0 and result['text_count']==1
    assert result['elements'][0]['observed_value']=='Label'


def test_numeric_text_inside_a_small_bordered_input_is_not_removed():
    image = Image.new('RGB',(120,80),'white')
    draw = ImageDraw.Draw(image)
    draw.rectangle((20,20,35,35),outline='black')
    draw.text((23,21),'0',fill='black')
    result = _combine(image,[(18,18,38,38)],_ocr([(22,21,33,34)],['0']))
    assert result['graphic_count']==1 and result['text_count']==1
    assert result['elements'][1]['observed_value']=='0'


def test_wide_light_bordered_input_and_button_survive_text_only_filter():
    image = Image.new('RGB', (520,160), 'white')
    draw = ImageDraw.Draw(image)
    draw.rectangle((20,20,460,46), outline=(192,192,192))
    draw.text((24,25), '1', fill='black')
    draw.rectangle((360,100,430,126), outline=(207,207,207))
    draw.text((375,106), 'OK', fill='black')
    result = _combine(image, [(19,19,448,48),(359,99,433,129)],
                      _ocr([(24,25,33,35),(375,106,392,116)], ['1','OK']))
    assert result['graphic_count'] == result['text_count'] == 2
    assert {i for e in result['elements'] for i in e.get('yolo_proposal_ids',[])} == {1,2}


def test_grid_cells_are_audited_as_layout_without_losing_checkbox_or_text():
    image = Image.new('RGB', (360,240), 'white')
    draw = ImageDraw.Draw(image)
    for y in (20,60,100,140,180):
        draw.line((20,y,320,y),fill=(190,190,190))
    for x in (20,120,220,320):
        draw.line((x,20,x,180),fill=(190,190,190))
    draw.text((35,72),'1',fill='black')
    draw.rectangle((150,70,165,85),outline='black')
    result = _combine(image, [(20,60,120,100),(148,68,168,88)],
                      _ocr([(35,72,44,82)],['1']))
    assert result['graphic_count'] == result['text_count'] == 1
    assert result['elements'][0]['yolo_proposal_ids']==[2]
    region = result['ui_fusion']['layout_regions'][0]
    assert region['id']==1 and region['layout_kind']=='grid-cell'
    assert result['ui_fusion']['grids']


def test_repeated_arrow_ocr_becomes_one_graphic_per_location():
    image = Image.new('RGB',(140,140),'white')
    draw = ImageDraw.Draw(image)
    for y in (20,60,100):
        draw.line([(20,y),(25,y+5),(20,y+10)],fill='black')
    ocr = _ocr([(17,18,29,33),(17,58,29,73),(17,98,29,113)],['>','>','>'])
    result = _combine(image,[(14,14,34,38)],ocr)
    assert result['graphic_count']==3 and result['text_count']==0
    assert len(result['ui_fusion']['rejected_ocr'])==3
    assert ocr['texts']==['>','>','>']


def test_symbol_text_is_retained_without_independent_icon_evidence():
    image = Image.new('RGB',(140,70),'white')
    ImageDraw.Draw(image).text((20,20),'>',fill='black')
    result = _combine(image,texts=_ocr([(18,18,32,34)],['>']))
    assert result['graphic_count']==0 and result['text_count']==1


def test_window_boundary_alone_is_not_a_control_but_close_glyph_is():
    image = Image.new('RGB',(200,120),'white')
    draw = ImageDraw.Draw(image)
    draw.line((0,0,0,119),fill='black')
    draw.line((174,10,185,21),fill='black')
    draw.line((185,10,174,21),fill='black')
    result = _combine(image,[(0,50,14,90),(170,6,190,26)],
                      _ocr([(169,5,191,27)],['×']))
    assert result['graphic_count']==1 and result['text_count']==0
    assert result['elements'][0]['geometry_kind']=='compact-icon'
    assert result['elements'][0]['yolo_proposal_ids']==[2]


def test_magnifier_can_be_split_from_cjk_ocr_without_changing_raw_ocr():
    image = Image.new('RGB',(220,100),'white')
    draw = ImageDraw.Draw(image)
    draw.text((20,25),'Name',fill='black')
    draw.ellipse((65,20,77,32),outline='black',width=2)
    draw.line((75,30,81,36),fill='black',width=2)
    ocr = _ocr([(19,19,83,38)],['信号名称Q'])
    result = _combine(image,[(15,15,90,42)],ocr)
    assert len(result['ui_fusion']['text_corrections'])==1
    assert any(e.get('geometry_kind')=='magnifier-icon' for e in result['elements'])
    text = next(e for e in result['elements'] if e['entity_type']=='text-label')
    assert text['observed_value']=='信号名称'
    assert text['ocr_layout_correction']['before']['text']=='信号名称Q'
    assert ocr['texts']==['信号名称Q']


def test_regular_ascii_q_is_not_split_as_a_search_icon():
    image = Image.new('RGB',(220,100),'white')
    ImageDraw.Draw(image).text((20,20),'QualityQ',fill='black')
    result = _combine(image,[(15,15,90,42)],_ocr([(19,19,82,35)],['QualityQ']))
    assert result['ui_fusion']['text_corrections']==[]
    assert result['elements'][-1]['observed_value']=='QualityQ'


def test_loose_yolo_box_does_not_promote_an_ocr_digit_into_an_icon():
    image = Image.new('RGB',(120,70),'white')
    ImageDraw.Draw(image).text((30,25),'7',fill='black')
    result = _combine(image,[(20,15,55,50)],_ocr([(29,24,38,37)],['7']))
    assert result['graphic_count']==0 and result['text_count']==1
    assert result['ui_fusion']['template_count']==0


def test_scroll_thumb_needs_both_filled_pixel_evidence_and_a_nearby_region():
    image = Image.new('RGB',(280,200),'white')
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((150,30,157,150),radius=3,fill=(216,216,216))
    draw.rounded_rectangle((230,30,237,150),radius=3,fill=(216,216,216))
    result = _combine(image,[(20,25,149,155)])
    details=result['ui_fusion']['geometry_additions']
    assert len(details)==1 and details[0]['geometry_kind']=='scrollbar-thumb'
    assert details[0]['rect']['left'] < 160


def test_low_contrast_toolbar_icons_are_not_deleted_as_low_score_ocr_letters():
    image=Image.new('RGB',(180,80),'white')
    draw=ImageDraw.Draw(image)
    for x in (20,60,100):
        draw.polygon([(x+3,23),(x+16,30),(x+3,37)],outline=(210,210,210))
    ocr=_ocr([(20,20,40,40),(60,20,80,40),(100,20,120,40)],['D','D','D'])
    ocr['scores']=[0.7]*3
    result=_combine(image,[(19,19,41,41),(59,19,81,41),(99,19,121,41)],ocr)
    assert result['graphic_count']==3 and result['text_count']==0
    assert result['ui_fusion']['compact_icon_rows'][0]['distinct_controls']==3
    assert all(e['rect']['width']>=12 for e in result['elements'])


def test_confident_numeric_text_rows_do_not_become_toolbar_icons():
    image=Image.new('RGB',(180,80),'white')
    draw=ImageDraw.Draw(image)
    for x,t in zip((20,60,100),('1','2','3')):
        draw.text((x,23),t,fill='black')
    result=_combine(image,[(18,18,40,40),(58,18,80,40),(98,18,120,40)],
                    _ocr([(20,20,30,38),(60,20,70,38),(100,20,110,38)],['1','2','3']))
    assert not result['ui_fusion']['compact_icon_rows']
    assert result['graphic_count']==0 and result['text_count']==3


def test_a_closed_character_inside_an_ocr_word_is_not_a_checkbox():
    image=Image.new('RGB',(180,80),'white')
    draw=ImageDraw.Draw(image)
    draw.text((20,20),'Word',fill='black')
    draw.rectangle((85,22,96,33),outline='black')
    result=_combine(image,[(83,20,99,36)],_ocr([(20,19,110,38)],['命令窗口']))
    assert result['graphic_count']==0 and result['text_count']==1


def test_partial_text_prediction_does_not_survive_as_a_graphic_fragment():
    image=Image.new('RGB',(160,80),'white')
    ImageDraw.Draw(image).text((20,20),'Long label',fill='black')
    result=_combine(image,[(18,18,40,37)],_ocr([(19,19,95,36)],['Long label']))
    assert result['graphic_count']==0 and result['text_count']==1


def test_repeated_tree_prefixes_are_separated_from_the_caption():
    image=Image.new('RGB',(240,180),'white')
    draw=ImageDraw.Draw(image)
    for y in (20,60,100):
        draw.line([(20,y+5),(25,y+10),(30,y+5)],fill='black')
        draw.rectangle((42,y,56,y+14),outline='black')
        draw.text((70,y),'Node',fill='black')
    result=_combine(image,[(15,y-3,140,y+20) for y in (20,60,100)],
        _ocr([(18,y-3,140,y+20) for y in (20,60,100)],['节点1','节点2','节点3']))
    kinds=[e.get('geometry_kind') for e in result['elements']]
    assert kinds.count('tree-chevron')==3 and kinds.count('tree-prefix-icon')==3
    texts=[e for e in result['elements'] if e['entity_type']=='text-label']
    assert all(t['rect']['left']>=68 for t in texts)
    assert [t['observed_value'] for t in texts]==['节点1','节点2','节点3']


def test_prefix_search_symbol_is_separated_without_rewriting_the_raw_ocr():
    image=Image.new('RGB',(180,90),'white')
    draw=ImageDraw.Draw(image)
    draw.ellipse((22,20,34,32),outline='black',width=2)
    draw.line((32,30,38,36),fill='black',width=2)
    ocr=_ocr([(19,18,130,39)],['Q搜索'])
    result=_combine(image,[(15,15,140,45)],ocr)
    assert any(e.get('geometry_kind')=='magnifier-icon' for e in result['elements'])
    assert next(e for e in result['elements'] if e['entity_type']=='text-label')['observed_value']=='搜索'
    assert ocr['texts']==['Q搜索']
