"""Exercise the GUI image fallback with real templates and no desktop input."""
from copy import deepcopy
from types import SimpleNamespace

import numpy as np
from PIL import Image
import pytest

from easy_uiauto import library_actions as actions, workflows
from easy_uiauto.locator_images import ImageTarget, reference_images
from easy_uiauto.mcp import control_observation
from easy_uiauto.mcp.protocol import location_from_xpath
from easy_uiauto.mcp.results import ActionError


@pytest.fixture
def visual(monkeypatch,tmp_path):
    import uiautomation
    template=Image.fromarray(np.random.default_rng(9).integers(0,256,(24,24,3),dtype=np.uint8))
    path=tmp_path/'button.png';template.save(path)
    screen=Image.new('RGB',(240,160),'white');screen.paste(template,(70,45))
    bounds=dict(left=-300,top=100,right=-60,bottom=260,width=240,height=160)
    root=SimpleNamespace(NativeWindowHandle=1,ProcessId=10,ClassName='AppClass')
    expected=dict(handle=1,process_id=10,class_name='AppClass')
    record=dict(id='button',name='Button',status='observed',image='button.png',
        location=location_from_xpath([dict(ControlType='WindowControl',Name='App',ClassName='AppClass'),
                                    dict(ControlType='ButtonControl',Name='Button')]))
    events=[]
    def missing(*a,**kw):raise ActionError('control_not_found','Control not found')
    monkeypatch.setattr(actions,'observe_control',missing)
    monkeypatch.setattr(actions,'bind_window',lambda _:root)
    monkeypatch.setattr(actions,'owned_windows',lambda _,**kw:[])
    monkeypatch.setattr(actions.native_windows,'enumerate_windows',lambda:[{**expected,'visible':True,'enabled':True,'owner_handle':0,'title':'App'}])
    monkeypatch.setattr(actions,'activate_application',lambda _:events.append('activate-app'))
    monkeypatch.setattr(actions.Controller,'activate_window_handle',lambda *a,**k:events.append('activate'))
    monkeypatch.setattr(actions.scanner,'_window_rect',lambda _:dict(bounds))
    monkeypatch.setattr(actions.scanner,'_screenshot_window',lambda _:screen.copy())
    monkeypatch.setattr(uiautomation,'GetForegroundWindow',lambda:1)
    monkeypatch.setattr(uiautomation,'IsIconic',lambda _:False)
    monkeypatch.setattr(actions.Controller,'execute_resolved',lambda action,**kw:events.append((action,kw)) or dict(executed=True))
    return SimpleNamespace(directory=tmp_path,record=record,expected=expected,root=root,screen=screen,
                           template=template,events=events,bounds=bounds)


def test_image_miss_uses_fresh_pixels_correct_screen_origin_and_clicks_once(visual):
    v=visual;original=deepcopy(v.record)
    result=actions.execute_plan([v.record],v.expected,directory=v.directory,
        before_click=lambda r,c,w:control_observation.capture(v.directory,c,w,phase='before-action')[0])
    sent=[item for item in v.events if isinstance(item,tuple)]
    assert len(sent)==1 and sent[0][0]=='click' and 'control' not in sent[0][1]
    assert sent[0][1]['rectangle']['left']==-230 and sent[0][1]['rectangle']['top']==145
    assert result['steps'][0]['resolved_by']=='image' and result['steps'][0]['before']['resolved_by']=='image'
    assert v.record==original


@pytest.mark.parametrize('code,reason',[
    ('ambiguous_match',''),('invalid_argument',''),('control_not_found','index_out_of_range'),
    ('control_not_found','position_context_missing'),('control_not_found','identity_verification_failed'),
    ('window_invalid',''),('operation_blocked','uia_access_denied')])
def test_terminal_locator_failures_never_activate_or_try_images(visual,monkeypatch,code,reason):
    v=visual
    def reject(*a,**kw):raise ActionError(code,'Rejected',reason=reason)
    monkeypatch.setattr(actions,'observe_control',reject)
    with pytest.raises(ActionError):actions.resolve_record(v.record,v.expected,v.root,directory=v.directory)
    assert v.events==[]


def test_position_mode_never_falls_back_to_another_picture(visual):
    v=visual;v.record['location']['PARAMETERS']['match_mode']='position'
    with pytest.raises(ActionError):actions.resolve_record(v.record,v.expected,v.root,directory=v.directory)
    assert v.events==[]


def test_missing_explicit_image_does_not_use_an_older_learned_template(visual):
    v=visual;v.record['location']['Img']='missing.png'
    assert reference_images(v.directory,v.record)==[]
    with pytest.raises(ActionError,match='没有可用'):actions.resolve_record(v.record,v.expected,v.root,directory=v.directory)
    assert v.events==[]


def test_duplicate_icons_stop_without_clicking(visual):
    v=visual;v.screen.paste(v.template,(150,45))
    with pytest.raises(ActionError,match='图片匹配未通过'):
        actions.execute_plan([v.record],v.expected,directory=v.directory)
    assert not any(isinstance(e,tuple) for e in v.events)


def test_target_changes_between_match_and_capture_stop_before_input(visual):
    v=visual
    def before(r,c,w):
        v.screen.paste('white',(70,45,94,69));v.screen.paste(v.template,(130,70))
        return control_observation.capture(v.directory,c,w,phase='before-action')[0]
    with pytest.raises(ActionError,match='截图时发生变化'):
        actions.execute_plan([v.record],v.expected,directory=v.directory,before_click=before)
    assert not any(isinstance(e,tuple) for e in v.events)


def test_explicit_visual_observation_replays_without_a_fabricated_xpath(visual):
    v=visual
    v.record['location']=location_from_xpath([],parameters={'locator_kind':'visual-observation'})
    v.record['location']['Img']='button.png'
    plan=actions.validate_plan([v.record],v.record['id'])
    assert actions.execute_plan(plan,v.expected,directory=v.directory)['ok']
    assert not v.record['location']['Xpath']


def test_workflow_keeps_image_fallback_for_later_steps_and_records_the_method(visual):
    v=visual
    flow=dict(keyword='Two clicks',steps=[workflows.click_step(v.record),workflows.click_step(v.record)])
    result=workflows.execute(flow,[v.record],v.expected,directory=v.directory)
    assert result['ok'] and [s['resolved_by'] for s in result['steps']]==['image','image']
    assert len([e for e in v.events if isinstance(e,tuple)])==2


def test_uncertain_semantics_do_not_block_user_selected_image_click(visual):
    v=visual
    v.record.update(status='quarantined',semantic_status='uncertain',
                    verification=dict(location='failed',image='passed'))
    flow=dict(keyword='Manual image click',steps=[workflows.click_step(v.record)])
    assert workflows.readiness(flow,[v.record])['ready']
    assert workflows.execute(flow,[v.record],v.expected,directory=v.directory)['ok']
    assert v.record['semantic_status']=='uncertain' and len([e for e in v.events if isinstance(e,tuple)])==1


@pytest.mark.parametrize('location',[None,{'Img':'button.png'}])
def test_image_only_legacy_records_need_no_uia_properties_or_kind_flag(visual,location):
    v=visual;v.record['location']=deepcopy(location);original=deepcopy(v.record)
    flow=dict(keyword='Image only',steps=[workflows.click_step(v.record)])
    assert workflows.readiness(flow,[v.record])['ready']
    result=workflows.execute(flow,[v.record],v.expected,directory=v.directory)
    assert result['ok'] and result['steps'][0]['resolved_by']=='image'
    assert v.record==original


def test_real_runtime_quarantine_still_stops_before_lookup_or_input(visual,monkeypatch):
    v=visual
    v.record.update(status='quarantined',quarantine=dict(active=True,reason='Failed runtime match'))
    flow=dict(keyword='Blocked stale record',steps=[workflows.click_step(v.record)])
    with pytest.raises(ValueError,match='隔离'):
        workflows.execute(flow,[v.record],v.expected,directory=v.directory)
    assert v.events==[]


def test_image_target_cannot_claim_uia_properties(visual):
    v=visual;target=actions.resolve_record(v.record,v.expected,v.root,directory=v.directory)
    assert isinstance(target,ImageTarget) and workflows._check(dict(kind='exists'),target)['exists']
    with pytest.raises(ActionError,match='无法读取'):
        workflows._check(dict(kind='read-only',expected=True),target)


def test_verify_worker_passes_the_image_record_and_returns_geometry(visual,monkeypatch):
    from easy_uiauto import library_tasks
    from easy_uiauto.knowledge_library import KnowledgeLibrary
    import json
    v=visual;folder=v.directory/'job';folder.mkdir()
    monkeypatch.setattr(KnowledgeLibrary,'directory',lambda *a:v.directory)
    library_tasks._write(folder/'request.json',dict(operation='verify',vault=str(v.directory/'vault'),
        app_id='demo',window=v.expected,**v.record))
    library_tasks._worker(folder)
    state=json.loads((folder/'state.json').read_text(encoding='utf-8'))
    assert state['state']=='completed' and state['result']['resolved_by']=='image'
    assert state['result']['rect']['left']==-230
    assert not any(isinstance(e,tuple) for e in v.events)


def test_foreground_permission_is_granted_to_only_the_requested_worker(monkeypatch):
    from easy_uiauto.mcp import native_windows
    calls=[]
    class Grant:
        def __call__(self,pid):calls.append(pid);return True
    monkeypatch.setattr(native_windows.ctypes,'WinDLL',lambda *a,**k:SimpleNamespace(AllowSetForegroundWindow=Grant()))
    assert native_windows.allow_foreground(4321) and calls==[4321]


def test_cancelled_gui_job_does_not_bind_or_activate_before_parent_is_ready(tmp_path,monkeypatch):
    from easy_uiauto import library_tasks
    import json
    library_tasks._write(tmp_path/'request.json',dict(operation='click',capture_handshake=True))
    (tmp_path/'cancel').touch()
    monkeypatch.setattr(library_tasks,'_run_worker',lambda _:pytest.fail('Worker started before GUI was ready'))
    library_tasks._worker(tmp_path)
    assert json.loads((tmp_path/'state.json').read_text(encoding='utf-8'))['state']=='cancelled'
