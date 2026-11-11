"""Detect disabled owners and manage only explicitly selected owned windows."""
from __future__ import annotations

import time
from .mcp import native_windows
from .mcp.results import ActionError


def inspect(expected, scope=None):
    from .library_actions import owned_windows
    windows=native_windows.enumerate_windows()
    by_handle={w['handle']:w for w in windows}
    main=by_handle.get(expected['handle'])
    if (not main or not main.get('visible') or
            any(main.get(k)!=expected.get(k) for k in ('process_id','class_name'))):
        raise ActionError('window_invalid','所选软件窗口已经失效，请重新选择')
    owned=owned_windows(expected,entries=windows)
    scope=scope or {'kind':'main'}
    target=main
    if scope.get('kind')=='owned':
        candidates=[w for w in owned if w['class_name']==scope.get('class_name') and
                    (not scope.get('title') or w['title']==scope['title'])]
        if not candidates:raise ActionError('control_not_found','目标弹出窗口尚未显示',retryable=True)
        if len(candidates)!=1:raise ActionError('ambiguous_match','多个弹窗符合目标上下文')
        target=candidates[0]
    def blocks(window):
        owner=window.get('owner_handle',0);seen=set()
        while owner and owner not in seen:
            if owner==target['handle']:return True
            seen.add(owner);owner=by_handle.get(owner,{}).get('owner_handle',0)
        return False
    blocked=not target.get('enabled',True)
    popups=[{**w,'blocking':blocked and w.get('enabled',True) and blocks(w)} for w in owned]
    popups.sort(key=lambda w:(not w['blocking'],not w.get('foreground'),w.get('z_order',0)))
    return dict(ok=True,main_window=main,target_window=target,blocked=blocked,popups=popups,
                message='目标窗口被弹窗阻挡，请先处理弹窗' if blocked else '未发现阻挡目标窗口的弹窗')


def assert_unblocked(expected,scope=None):
    # Old internal callers may supply only a handle; GUI bindings always include identity.
    if not all(k in expected for k in ('handle','process_id','class_name')):return
    state=inspect(expected,scope)
    if state['blocked']:
        names=[p['title'] or p['class_name'] for p in state['popups'] if p['blocking']]
        detail='、'.join(names[:4]) if names else '未找到可确认所属的弹窗，目标窗口当前已禁用'
        error=ActionError('operation_blocked','目标窗口被阻挡：'+detail+'。请先处理弹窗后重新执行。',reason='modal_blocked')
        error.popup_state=state
        raise error


def selected_popup(expected,selection):
    state=inspect(expected)
    selected=next((w for w in state['popups'] if w['handle']==selection.get('handle')),None)
    if (not selected or any(selected.get(k)!=selection.get(k) for k in
                            ('process_id','class_name','title','owner_handle'))):
        raise ActionError('window_invalid','所选弹窗已关闭或发生变化，请刷新列表')
    if not selected.get('enabled',True):
        error=ActionError('operation_blocked','此弹窗也被其他弹窗阻挡，请先处理最内层弹窗',reason='modal_blocked')
        error.popup_state=state;raise error
    return selected


def focus(expected,selection):
    from .ctrl import Controller
    popup=selected_popup(expected,selection)
    Controller.activate_window_handle(popup['handle'])
    return {**inspect(expected),'message':'已激活弹窗：'+(popup['title'] or popup['class_name'])}


def close(expected,selection):
    """Send one normal WM_CLOSE; never confirm, discard, kill, or retry an input."""
    popup=selected_popup(expected,selection)
    if not native_windows.request_close_window(popup['handle']):
        raise ActionError('action_failed','Windows 未接受弹窗关闭请求；未发送确认或取消按键')
    deadline=time.monotonic()+.8
    while True:
        state=inspect(expected)
        closed=not any(w['handle']==popup['handle'] for w in state['popups'])
        if closed or time.monotonic()>=deadline:break
        time.sleep(.05)
    return {**state,'closed':closed,'input_attempted':True,
            'message':'所选弹窗已关闭；请重新执行原动作' if closed else
                      '已请求关闭；窗口仍存在，请查看是否需要处理保存确认或其他提示'}
