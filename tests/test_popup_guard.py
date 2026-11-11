from copy import deepcopy
from types import SimpleNamespace

import pytest

from easy_uiauto import popup_guard,library_actions,workflows
from easy_uiauto.mcp import native_windows
from easy_uiauto.mcp.results import ActionError
from easy_uiauto.mcp.protocol import location_from_xpath


@pytest.fixture
def windows(monkeypatch):
    expected=dict(handle=1,process_id=10,class_name='App')
    def entry(handle,owner=0,enabled=True,title='Dialog',pid=10,visible=True):
        return dict(handle=handle,owner_handle=owner,enabled=enabled,title=title,process_id=pid,
                    class_name='App' if handle==1 else 'Dialog',visible=visible,foreground=False,z_order=handle)
    entries=[entry(1,enabled=False,title='Application'),entry(2,owner=1)]
    monkeypatch.setattr(native_windows,'enumerate_windows',lambda:deepcopy(entries))
    return SimpleNamespace(expected=expected,entries=entries,entry=entry)


def test_modal_reports_owned_blocker_and_never_sends_input(windows,monkeypatch):
    v=windows;calls=[]
    monkeypatch.setattr(library_actions.Controller,'activate_window_handle',lambda *a,**k:calls.append('focus'))
    monkeypatch.setattr(library_actions.Controller,'execute_resolved',lambda *a,**k:calls.append('input'))
    with pytest.raises(ActionError) as failure:popup_guard.assert_unblocked(v.expected)
    assert failure.value.reason=='modal_blocked' and failure.value.popup_state['popups'][0]['title']=='Dialog'
    assert calls==[]


def test_modeless_tool_window_is_listed_without_blocking(windows):
    windows.entries[0]['enabled']=True
    state=popup_guard.inspect(windows.expected)
    assert not state['blocked'] and len(state['popups'])==1 and not state['popups'][0]['blocking']
    popup_guard.assert_unblocked(windows.expected)


def test_nested_modal_allows_only_the_innermost_target(windows):
    v=windows;v.entries[1]['enabled']=False;v.entries.append(v.entry(3,owner=2,title='Save confirmation'))
    assert popup_guard.inspect(v.expected)['popups'][0]['handle']==3
    with pytest.raises(ActionError):popup_guard.assert_unblocked(v.expected,dict(kind='owned',class_name='Dialog',title='Dialog'))
    popup_guard.assert_unblocked(v.expected,dict(kind='owned',class_name='Dialog',title='Save confirmation'))


def test_hidden_foreign_and_unrelated_windows_are_never_managed(windows):
    v=windows;v.entries[0]['enabled']=True
    v.entries.extend([v.entry(3,owner=1,pid=20),v.entry(4,owner=0),v.entry(5,owner=1,visible=False)])
    assert [p['handle'] for p in popup_guard.inspect(v.expected)['popups']]==[2]


@pytest.mark.parametrize('mutation',['handle','process_id','class_name','title','owner_handle'])
def test_stale_or_foreign_selection_cannot_close_any_window(windows,monkeypatch,mutation):
    v=windows;selected=deepcopy(v.entries[1]);selected[mutation]='wrong' if mutation in {'class_name','title'} else 999
    calls=[];monkeypatch.setattr(native_windows,'request_close_window',lambda h:calls.append(h))
    with pytest.raises(ActionError):popup_guard.close(v.expected,selected)
    assert calls==[]


def test_close_requests_only_the_selected_popup_once_and_does_not_retry_original_action(windows,monkeypatch):
    v=windows;calls=[]
    def close(handle):
        calls.append(handle);v.entries.pop();v.entries[0]['enabled']=True;return True
    monkeypatch.setattr(native_windows,'request_close_window',close)
    result=popup_guard.close(v.expected,deepcopy(v.entries[1]))
    assert calls==[2] and result['closed'] and not result['blocked'] and result['input_attempted']


def test_close_that_opens_a_save_confirmation_stops_without_accepting_it(windows,monkeypatch):
    v=windows;calls=[]
    def close(handle):
        calls.append(handle);v.entries[1]['enabled']=False
        v.entries.append(v.entry(3,owner=2,title='Save changes?'));return True
    monkeypatch.setattr(native_windows,'request_close_window',close)
    monkeypatch.setattr(popup_guard.time,'monotonic',iter([0.,1.]).__next__)
    state=popup_guard.close(v.expected,deepcopy(v.entries[1]))
    assert calls==[2] and not state['closed'] and state['blocked']
    assert state['popups'][0]['title']=='Save changes?'


def test_main_window_is_never_a_close_target(windows):
    with pytest.raises(ActionError):popup_guard.selected_popup(windows.expected,windows.entries[0])


def test_modal_guard_runs_before_uia_or_image_fallback(windows,monkeypatch,tmp_path):
    def forbidden(*a,**k):pytest.fail('Lookup or image fallback ran behind a modal')
    monkeypatch.setattr(library_actions,'observe_control',forbidden)
    monkeypatch.setattr(library_actions,'resolve_image',forbidden)
    record=dict(location=location_from_xpath([dict(ControlType='WindowControl',Name='App'),dict(ControlType='ButtonControl',Name='Button')]))
    with pytest.raises(ActionError,match='阻挡'):
        library_actions.resolve_record(record,windows.expected,object(),directory=tmp_path)


def test_modal_appearing_during_capture_stops_workflow_before_input(windows,monkeypatch):
    import uiautomation
    v=windows;v.entries[0]['enabled']=True;v.entries.pop();calls=[]
    monkeypatch.setattr(library_actions,'bind_window',lambda _:object())
    monkeypatch.setattr(library_actions,'activate_application',lambda _:None)
    monkeypatch.setattr(library_actions,'resolve_record',lambda *a,**k:object())
    monkeypatch.setattr(uiautomation,'IsIconic',lambda _:False)
    monkeypatch.setattr(uiautomation,'GetForegroundWindow',lambda:1)
    monkeypatch.setattr(library_actions.Controller,'execute_resolved',lambda *a,**k:calls.append('click'))
    record=dict(id='button',name='Button',status='observed',location=location_from_xpath([
        dict(ControlType='WindowControl',Name='App'),dict(ControlType='ButtonControl',Name='Button')]))
    def capture(*a):
        v.entries[0]['enabled']=False;v.entries.append(v.entry(2,owner=1));return {}
    with pytest.raises(ActionError) as failure:
        workflows.execute(dict(keyword='Click',steps=[workflows.click_step(record)]),[record],v.expected,before_action=capture)
    assert not calls and not failure.value.execution_result['input_attempted']
    assert failure.value.execution_result['popup_state']['blocked']


def test_native_close_is_one_asynchronous_wm_close(monkeypatch):
    calls=[]
    class Post:
        def __call__(self,*args):calls.append(args);return True
    monkeypatch.setattr(native_windows.ctypes,'WinDLL',lambda *a,**k:SimpleNamespace(PostMessageW=Post()))
    assert native_windows.request_close_window(2) and calls==[(2,0x0010,0,0)]


def test_worker_returns_blocking_window_metadata_without_touching_uia(windows,tmp_path):
    import json
    from easy_uiauto import library_tasks
    library_tasks._write(tmp_path/'request.json',dict(operation='inspect-popups',window=windows.expected))
    library_tasks._worker(tmp_path)
    result=json.loads((tmp_path/'state.json').read_text(encoding='utf-8'))['result']
    assert result['blocked'] and result['popups'][0]['handle']==2


def test_popup_manager_does_not_close_automatically_and_buttons_target_selection(windows):
    pytest.importorskip('PySide6')
    from PySide6.QtWidgets import QApplication
    from easy_uiauto.qt_ui.popup_dialog import PopupDialog
    app=QApplication.instance() or QApplication([]);calls=[]
    dialog=PopupDialog(popup_guard.inspect(windows.expected),lambda *args:calls.append(args))
    try:
        assert not calls and dialog.close_button.isEnabled()
        dialog.close_button.click()
        assert len(calls)==1 and calls[0][0]=='close-popup' and calls[0][1]['handle']==2
        dialog.set_state(dict(popups=[],message='没有弹窗'))
        assert not dialog.close_button.isEnabled() and not dialog.focus_button.isEnabled()
    finally:dialog.close();app.processEvents()
