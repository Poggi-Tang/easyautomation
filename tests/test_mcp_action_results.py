"""Portable contract tests; fake desktop callbacks do not validate Windows UIA.

Load the pure result module and the actual server function definitions in an
isolated namespace. This avoids importing native desktop libraries on Linux while
still exercising the production wrappers and FastMCP registration/schema.
"""
from __future__ import annotations

import ast
import asyncio
import importlib
import sys
from contextlib import nullcontext
from functools import wraps
import importlib.util
import json
from pathlib import Path
import time
from types import ModuleType, SimpleNamespace
from typing import Optional
from unittest.mock import Mock

import pytest

ROOT = Path(__file__).resolve().parents[1] / "src" / "easy_uiauto" / "mcp"
spec = importlib.util.spec_from_file_location("_action_result_contract", ROOT / "results.py")
results = importlib.util.module_from_spec(spec)
spec.loader.exec_module(results)
package = ModuleType('_easy_action_contract')
package.__path__ = [str(ROOT)]
sys.modules[package.__name__] = package
sys.modules[package.__name__ + '.results'] = results
waiting = importlib.import_module(package.__name__ + '.waiting')
protocol = importlib.import_module(package.__name__ + '.protocol')


@pytest.mark.parametrize("value,code", [
    ('{"ok":false,"error":"boom"}', 'action_failed'),
    ('{"found":false}', 'control_not_found'),
    ('Window not found or could not be activated: Editor', 'window_invalid'),
    ('Error: [ambiguous_match] Multiple matches', 'ambiguous_match'),
    ('步骤：点击 异常：[window_invalid] Window disappeared', 'window_invalid'),
    ('Error: Text paste could not be read back', 'result_unconfirmed'),
    ('Error: Current mode is learn; mutating operations are blocked', 'operation_blocked'),
])
def test_failures_are_machine_readable(value, code):
    outcome = results.normalize_result(value)
    assert outcome['ok'] is False
    assert outcome['code'] == code
    assert outcome['verified'] is False
    assert outcome['retryable'] is False


def test_action_execution_is_not_verification():
    outcome = results.normalize_result('Clicked left at (10, 20)')
    assert outcome['ok'] is True
    assert outcome['verified'] is False
    assert outcome['code'] == 'action_executed'
    assert results.normalize_result(outcome) == outcome


def test_readback_evidence_and_original_fields_are_preserved():
    outcome = results.normalize_result({
        'ok': True, 'command': 'editor.set-text',
        'action_verification': {'verified': True, 'evidence': 'accessible-value'},
    })
    assert outcome['verified'] is True
    assert outcome['command'] == 'editor.set-text'
    assert outcome['data']['action_verification']['evidence'] == 'accessible-value'


@pytest.fixture
def server():
    from mcp.server.fastmcp import FastMCP
    names = {
        '_action_result', '_is_error_result', '_result_ok', '_mode_blocks_operation',
        '_normalize_batch_step', '_require_active_window', '_recorded_action_is_mutating',
        'scroll', 'mouse_scroll', 'activate_window', 'press_key', 'hotkey',
        'run_action', 'run_actions', 'get_window_state', '_observe_window', '_rect_to_dict',
        'wait_for_control', 'wait_for_text', 'wait_for_window', '_observe_control',
        '_validate_click', 'click_at_position',
    }
    tree = ast.parse((ROOT / 'server.py').read_text(encoding='utf-8'))
    nodes = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names]
    namespace = {
        'json': json, 'time': time, 'wraps': wraps, 'Optional': Optional,
        'waiting': waiting, 'normalize_location': protocol.normalize_location,
        'native_windows': SimpleNamespace(matching_windows=Mock(return_value=[])),
        'observe_control': Mock(return_value=SimpleNamespace(IsEnabled=True)),
        'ui_cli': SimpleNamespace(_read_control_text=Mock(return_value=(True, 'Saved'))),
        '_auto_learn_point': Mock(),
        'mcp': FastMCP('contract-test'), 'strict_control_resolution': lambda **kwargs: nullcontext(),
        'normalize_result': results.normalize_result, 'action_result': results.result,
        'ActionError': results.ActionError, 'CURRENT_MODE': 'operate',
        'MUTATING_METHODS': {'click', 'scroll', 'mouse_scroll', 'press_key', 'run_action'},
        'MUTATING_RECORDED_ACTIONS': {'点击'},
        'pyautogui': SimpleNamespace(scroll=Mock(), moveTo=Mock(), click=Mock(), doubleClick=Mock()),
        'Controller': SimpleNamespace(key_click=Mock(return_value='Success'), key_group=Mock(return_value='Success')),
        'set_top_window': Mock(return_value=True), '_auto_learn_recorded_action': Mock(),
        '_run_action_easy': Mock(return_value='Success'),
    }
    for name in ('get_mode', 'set_mode', 'click', 'click_at_position', 'drag_control',
                 'find_control', 'get_control_at_position', 'get_control_tree',
                 'list_windows', 'move_mouse', 'set_text', 'type_text', 'wait_for_control',
                 'wait_for_text', 'wait_for_window'):
        namespace[name] = Mock(return_value='Success')
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(ROOT / 'server.py'), 'exec'), namespace)
    return namespace


def test_scroll_alias_uses_same_direction_and_coordinates(server):
    for name in ('scroll', 'mouse_scroll'):
        outcome = json.loads(server[name](amount=2, direction='up', x=-500, y=40))
        assert outcome['ok'] is True
        server['pyautogui'].scroll.assert_called_with(2)
        server['pyautogui'].moveTo.assert_called_with(-500, 40)
        server[name](amount=3, direction='down')
        server['pyautogui'].scroll.assert_called_with(-3)


@pytest.mark.parametrize('arguments', [
    {'amount': -1}, {'amount': 0}, {'amount': True}, {'direction': 'sideways'}, {'x': 40},
])
def test_invalid_scroll_does_not_move_or_scroll(server, arguments):
    outcome = json.loads(server['scroll'](**arguments))
    assert outcome['code'] == 'invalid_argument'
    server['pyautogui'].scroll.assert_not_called()
    server['pyautogui'].moveTo.assert_not_called()


def test_scroll_backend_failure_and_learn_mode(server):
    server['pyautogui'].scroll.side_effect = RuntimeError('device failed')
    assert json.loads(server['scroll']())['code'] == 'action_failed'
    server['CURRENT_MODE'] = 'learn'
    server['pyautogui'].scroll.reset_mock()
    assert json.loads(server['mouse_scroll']())['code'] == 'operation_blocked'
    server['pyautogui'].scroll.assert_not_called()


@pytest.mark.parametrize('name,arguments', [('press_key', {'key': 'enter'}), ('hotkey', {'keys': 'ctrl+s'})])
def test_failed_activation_never_sends_keys(server, name, arguments):
    server['set_top_window'].return_value = False
    outcome = json.loads(server[name](window_name='Gone', **arguments))
    assert outcome['code'] == 'window_invalid'
    server['Controller'].key_click.assert_not_called()
    server['Controller'].key_group.assert_not_called()


def test_batch_stops_on_structured_failure(server):
    server['click'].return_value = json.dumps({'ok': False, 'error': '[ambiguous_match] duplicate'})
    outcome = json.loads(server['run_actions'](json.dumps([
        {'method': 'click'}, {'method': 'press_key', 'params': {'key': 'enter'}},
    ])))
    assert outcome['ok'] is False
    assert outcome['executed'] == 1
    assert outcome['results'][0]['code'] == 'ambiguous_match'
    server['Controller'].key_click.assert_not_called()


def test_batch_non_object_failure_is_reported(server):
    outcome = json.loads(server['run_actions']('[42]'))
    assert outcome['ok'] is False
    assert outcome['results'][0]['code'] == 'invalid_argument'


def test_recorded_batch_calls_same_guarded_handler(server):
    server['CURRENT_MODE'] = 'learn'
    outcome = json.loads(server['run_actions'](json.dumps([
        {'method': 'run_action', 'params': {'action_json': json.dumps({'ACTION': '点击'})}},
    ])))
    assert outcome['ok'] is False
    server['_run_action_easy'].assert_not_called()


def test_fastmcp_keeps_action_schema(server):
    scroll = server['mcp']._tool_manager.get_tool('scroll')
    alias = server['mcp']._tool_manager.get_tool('mouse_scroll')
    assert set(scroll.parameters['properties']) == {'amount', 'direction', 'x', 'y'}
    assert scroll.parameters['properties'] == alias.parameters['properties']
    assert 'verified' in scroll.description


def test_wait_tools_integrate_with_result_protocol_and_uia_adapters(server):
    control = json.loads(server['wait_for_control']({'AutomationId': 'save'}, state='enabled', timeout_seconds=0))
    text = json.loads(server['wait_for_text']({'AutomationId': 'status'}, expected_text='Saved', timeout_seconds=0))
    assert control['verified'] is True
    assert text['verified'] is True
    assert text['data']['observation']['text'] == 'Saved'
    server['pyautogui'].click.assert_not_called()
    server['pyautogui'].moveTo.assert_not_called()


def test_wait_never_treats_invalid_window_as_disappearance(server):
    server['observe_control'].side_effect = results.ActionError('window_invalid', 'gone')
    value = json.loads(server['wait_for_control']({'AutomationId': 'save'}, state='absent', timeout_seconds=0))
    assert value['ok'] is False
    assert value['code'] == 'window_invalid'


def test_window_change_tracks_handle_after_title_changes(server):
    server['native_windows'].matching_windows.return_value = [dict(
        handle=42, title='Saved Editor', class_name='Editor', process_id=7,
        enabled=True, visible=True, bounds=None)]
    value = json.loads(server['wait_for_window']('Editor', state='changed',
                       baseline={'handle': 42, 'title': 'Editor'}, timeout_seconds=0))
    assert value['verified'] is True
    assert value['data']['observation']['title'] == 'Saved Editor'
    server['native_windows'].matching_windows.assert_called_once_with('Editor', handle=42)


def test_window_observation_rejects_ambiguous_native_title(server):
    server['native_windows'].matching_windows.return_value = [{'handle': 1}, {'handle': 2}]
    value = json.loads(server['get_window_state']('Editor'))
    assert value['ok'] is False and value['code'] == 'ambiguous_match'


def test_window_inventory_failure_is_not_treated_as_disappearance(server):
    server['native_windows'].matching_windows.side_effect = OSError('EnumWindows failed')
    value = json.loads(server['wait_for_window']('Editor', state='absent', timeout_seconds=0))
    assert value['ok'] is False and value['code'] == 'window_invalid'


def test_fastmcp_executes_wrapped_action_with_typed_parameters(server):
    tool = server['mcp']._tool_manager.get_tool('scroll')
    output = asyncio.run(tool.run({'amount': 2, 'direction': 'down'}))
    assert json.loads(output)['ok'] is True
    server['pyautogui'].scroll.assert_called_once_with(-2)


def test_invalid_click_choice_does_not_mutate(server):
    value = json.loads(server['click_at_position'](1, 2, button='unknown'))
    assert value['code'] == 'invalid_argument'
    server['pyautogui'].click.assert_not_called()


def test_legacy_success_with_error_words_in_control_name_is_not_failure():
    value = results.normalize_result('步骤：点击 控件【未找到异常失败】 执行动作 点击 成功！')
    assert value['ok'] is True


def test_partial_ui_command_batch_reports_actual_failure():
    value = results.normalize_result({'ok': False, 'completed_steps': 1,
        'failed_step': {'error': '[result_unconfirmed] Text could not be verified'},
        'steps': [{'command': 'click'}]})
    assert value['code'] == 'result_unconfirmed'
    assert value['completed_steps'] == 1
    assert value['retryable'] is False


def test_missing_match_does_not_report_completed_action():
    value = results.normalize_result({'found': False}, observation=True)
    assert value['ok'] is False
    assert value['code'] == 'control_not_found'
    assert value['message'] == 'No matching control was found'
