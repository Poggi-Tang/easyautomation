"""Native detail recovery must preserve scale, identity and the inference budget."""

import sys
from types import SimpleNamespace

from PIL import Image,ImageDraw

from easy_uiauto import image_details,image_detection


def test_native_small_inputs_and_blank_large_inputs_do_not_need_a_detail_pass():
    assert not image_details.plan(Image.new('RGB',(900,600),'white'))['enabled']
    large=Image.new('RGB',(2000,1400),'white')
    assert not image_details.plan(large)['enabled']


def test_content_regions_are_copied_without_resizing_and_fit_one_atlas():
    image=Image.new('RGB',(2200,1400),'white')
    draw=ImageDraw.Draw(image)
    for x,y in ((30,40),(2000,100),(400,1250)):
        draw.rectangle((x,y,x+18,y+18),outline='black',width=2)
    report=image_details.plan(image)
    assert report['enabled'] and report['parameters']['max_passes']==1
    atlas=image_details.render_atlas(image,report)
    for r in report['regions']:
        assert atlas.crop(tuple(r['atlas_rect'])).tobytes()==image.crop(tuple(r['source_rect'])).tobytes()
        assert 0<=r['atlas_rect'][0]<r['atlas_rect'][2]<=1280
        assert 0<=r['atlas_rect'][1]<r['atlas_rect'][3]<=1280


def _report():
    return dict(enabled=True,atlas_size=1280,source_size=[2000,1400],
                regions=[dict(id=1,source_rect=[100,200,164,264],atlas_rect=[16,16,80,80])])


def test_patch_coordinates_and_duplicates_do_not_merge_neighboring_controls():
    report=_report()
    original=[[103,203,121,221]]
    boxes,scores=image_details.merge_predictions(original,[0.8],
        [[20,20,36,36],[42,20,58,36],[72,20,88,36],[16,40,26,50]],[0.9,0.7,0.6,0.6],report)
    assert boxes==[[103,203,121,221],[126,204,142,220]]
    assert original==[[103,203,121,221]] and scores==[0.8,0.7]
    assert report['added_count']==report['duplicate_count']==1
    assert report['rejected_count']==2
    assert report['observations'][0]['merged_into_yolo_id']==1


def test_detector_loads_weights_once_for_whole_and_detail_passes(monkeypatch):
    loads=[]
    calls=[]
    model=SimpleNamespace(eval=lambda:model)
    monkeypatch.setitem(sys.modules,'torch',SimpleNamespace(
        set_num_threads=lambda _n:None,
        jit=SimpleNamespace(load=lambda *a,**kw:loads.append(a) or model)))
    def predict(image,got_model,confidence,**_options):
        calls.append(image.size)
        assert got_model is model and confidence==0.1
        return dict(boxes=[[103,203,121,221]] if len(calls)==1 else [[42,20,58,36]],
                    scores=[0.8],forward_seconds=0.2,elapsed_seconds=0.2)
    monkeypatch.setattr(image_detection,'_predict_yolo',predict)
    monkeypatch.setattr(image_details,'plan',lambda *_a:_report())
    result=image_detection.run_yolo(Image.new('RGB',(2000,1400),'white'),'weights',recover_details=True)
    assert len(loads)==1 and calls==[(2000,1400),(1280,1280)]
    assert result['inference_passes']==2 and result['whole_image_count']==1
    assert result['detail_recovery']['added_count']==1
    assert result['boxes'][0]==[103,203,121,221]
