"""Reduced padding and shared thread budgets must retain scale and error behavior."""

import threading

import pytest

from easy_uiauto import image_detection


@pytest.mark.parametrize('source,resized,canvas', [
    ((2560,1400),(1280,700),(1280,704)),
    ((1400,2560),(700,1280),(704,1280)),
    ((1280,1280),(1280,1280),(1280,1280)),
    ((221,79),(1280,458),(1280,480)),
])
def test_rectangular_inputs_keep_scale_and_align_padding_to_model_stride(source,resized,canvas):
    assert image_detection._yolo_input_size(source,rectangular=True)==(resized,canvas)
    assert image_detection._yolo_input_size(source)==(resized,(1280,1280))
    assert canvas[0]>=resized[0] and canvas[1]>=resized[1]
    assert all(side%32==0 for side in canvas)


def test_native_atlas_can_omit_unused_bottom_without_resizing_its_patches():
    assert image_detection._yolo_input_size((1280,1280),rectangular=True,
                                            content_bottom=604)==((1280,1280),(1280,608))
    assert image_detection._yolo_input_size((1280,1280),content_bottom=604)==(
        (1280,1280),(1280,1280))


def test_parallel_engines_share_thread_budget_and_report_on_the_caller_thread(monkeypatch):
    caller=threading.get_ident()
    rendezvous=threading.Barrier(2,timeout=5)
    calls={}
    messages=[]

    def run(name,*_args,**options):
        calls[name]=(threading.get_ident(),options)
        options['progress'](name+' progress')
        rendezvous.wait()
        return {'engine':name}

    monkeypatch.setattr(image_detection,'run_yolo',lambda *a,**kw:run('yolo',*a,**kw))
    monkeypatch.setattr(image_detection,'run_ocr',lambda *a,**kw:run('ocr',*a,**kw))
    yolo,ocr,execution=image_detection.run_detectors(
        object(),'weights',threads=4,recover_details=True,rectangular=True,
        progress=lambda message:messages.append((threading.get_ident(),message)))
    assert yolo['engine']=='yolo' and ocr['engine']=='ocr'
    assert all(ident!=caller for ident,_options in calls.values())
    assert calls['yolo'][1]['threads']==3 and calls['ocr'][1]['threads']==1
    assert calls['yolo'][1]['recover_details'] and calls['yolo'][1]['rectangular']
    assert execution['parallel'] and execution['thread_budget']==4
    assert all(ident==caller for ident,_message in messages)
    assert {'yolo progress','ocr progress'}<={message for _ident,message in messages}


@pytest.mark.parametrize('threads,parallel',[(1,True),(4,False)])
def test_single_thread_or_explicit_opt_out_retains_sequential_order(monkeypatch,threads,parallel):
    caller=threading.get_ident()
    calls=[]

    def run(name,*_args,**options):
        assert threading.get_ident()==caller and options['threads']==threads
        if name=='yolo':
            assert options['rectangular'] is False
        calls.append(name)
        return {'engine':name}

    monkeypatch.setattr(image_detection,'run_yolo',lambda *a,**kw:run('yolo',*a,**kw))
    monkeypatch.setattr(image_detection,'run_ocr',lambda *a,**kw:run('ocr',*a,**kw))
    _yolo,_ocr,execution=image_detection.run_detectors(object(),'weights',threads=threads,parallel=parallel)
    assert calls==['yolo','ocr'] and not execution['parallel']


def test_worker_failure_propagates_without_retries_or_leftover_workers(monkeypatch):
    rendezvous=threading.Barrier(2,timeout=5)
    finished=threading.Event()
    calls=[]

    def broken(*_a,**_kw):
        calls.append('yolo')
        rendezvous.wait()
        raise RuntimeError('failed local inference')

    def complete(*_a,**_kw):
        calls.append('ocr')
        rendezvous.wait()
        finished.set()
        return {}

    monkeypatch.setattr(image_detection,'run_yolo',broken)
    monkeypatch.setattr(image_detection,'run_ocr',complete)
    with pytest.raises(RuntimeError,match='failed local inference'):
        image_detection.run_detectors(object(),'weights')
    assert finished.is_set() and sorted(calls)==['ocr','yolo']
