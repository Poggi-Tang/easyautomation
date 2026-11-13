"""Regressions for the 2026-11-13 report, without desktop input or model calls."""

from contextlib import nullcontext
from copy import deepcopy
from pathlib import Path
from unittest.mock import Mock

import pytest
from PIL import Image, ImageDraw

from easy_uiauto.mcp import (interaction_learning, knowledge, library_tools, manual_layout,
                            native_windows, page_identity, scanner, ui_cli, window_identity)
from easy_uiauto.mcp.results import ActionError
from test_manual_region_scan import setup
from test_mcp_interaction_learning import _record
from test_mcp_scanner import FakeControl
from test_window_identity import app


@pytest.mark.parametrize('preserve_name', [False, True])
def test_initialization_preserves_binding_custom_metadata_and_body(tmp_path, monkeypatch, preserve_name):
    directory = knowledge.initialize_app('example', 'Human name', tmp_path)
    record, _ = knowledge.read_markdown(directory / 'app.md')
    record.update(window_name='Example', window_class='QtWindow', layout_source='manual',
                  description='Human description', custom={'nested': [1, 2]})
    knowledge.write_markdown(directory / 'app.md', record, '# Notes\n\nDo not erase.')
    knowledge.initialize_app('example', 'Scan name', tmp_path, preserve_name=preserve_name)
    updated, body = knowledge.read_markdown(directory / 'app.md')
    assert updated['name'] == ('Human name' if preserve_name else 'Scan name')
    for field in ('window_name', 'window_class', 'layout_source', 'description', 'custom'):
        assert updated[field] == record[field]
    assert 'Do not erase.' in body
    entry = dict(handle=7, title='Example', class_name='QtWindow', visible=True)
    monkeypatch.setattr(native_windows, 'enumerate_windows', lambda **kw: [entry])
    assert library_tools._window(updated, 0) == library_tools._window(updated, 7) == entry


def _page(identity=None, **changes):
    return dict(id='main', name='Main', window_identity=identity or dict(
        title='Main', class_name='QtWindow', uia_name='main', uia_class='Root', automation_id=''),
        layout_fingerprint='layout-main', **changes)


def _persist(directory, page, image):
    path = directory / 'images/pages' / f"{page['id']}.png"
    image.save(path)
    knowledge.save_page(directory, {**page, 'window_name': page['window_identity']['title'],
                                     'image': path.relative_to(directory).as_posix()})


@pytest.mark.parametrize('difference', ['title', 'class_name', 'uia_name', 'uia_class', 'size', 'layout'])
def test_white_dialogs_cannot_merge_across_window_or_page_identity(tmp_path, difference):
    directory = knowledge.initialize_app('example', 'Example', tmp_path)
    page = _page()
    image = Image.new('RGB', (720, 490), 'white')
    _persist(directory, page, image)
    other = deepcopy(page)
    if difference == 'size':
        image = Image.new('RGB', (1920, 1030), 'white')
    elif difference == 'layout':
        other['layout_fingerprint'] = 'other-controls'
        ImageDraw.Draw(image).text((10, 10), 'Different controls', fill='black')
    else:
        other['window_identity'][difference] = 'Different dialog'
    result = scanner._reuse_page_identity(directory, image, other)
    assert result['id'] != 'main'
    assert result['identity_decision'] == 'new-page-id-collision'
    assert (directory / 'images/pages/main.png').is_file()


def test_identical_page_reuses_id_even_when_model_renames_it(tmp_path):
    directory = knowledge.initialize_app('example', 'Example', tmp_path)
    image = Image.new('RGB', (720, 490), 'white')
    _persist(directory, _page(), image)
    result = scanner._reuse_page_identity(directory, image, {**_page(), 'id': 'renamed-by-ai'})
    assert result['id'] == 'main'
    assert result['previous_ai_id'] == 'renamed-by-ai'


def test_long_model_id_collision_remains_distinct_after_filename_slugging(tmp_path):
    directory = knowledge.initialize_app('example', 'Example', tmp_path)
    image = Image.new('RGB', (80, 60), 'white')
    page = {**_page(), 'id': 'a'*80}
    _persist(directory, page, image)
    other = {**page, 'window_identity': {'title': 'Other', 'uia_name': 'other'}}
    result = scanner._reuse_page_identity(directory, image, other)
    assert knowledge.slugify(result['id']) != knowledge.slugify(page['id'])
    _persist(directory, result, image)
    assert len(list(knowledge.iter_records(directory, 'page'))) == 2


def test_sparse_different_text_is_not_same_page_despite_high_mean_similarity():
    first, second = (Image.new('RGB', (720, 490), 'white') for _ in range(2))
    ImageDraw.Draw(first).text((30, 30), 'Open project', fill='black')
    ImageDraw.Draw(second).text((30, 30), 'Import package', fill='black')
    assert scanner._image_similarity(first, second) > .94
    assert page_identity.visual_match(first, second)[0] is False


def test_legacy_page_requires_matching_window_and_exact_image(tmp_path):
    directory = knowledge.initialize_app('example', 'Example', tmp_path)
    image = Image.new('RGB', (720, 490), 'white')
    image.save(directory / 'images/pages/main.png')
    knowledge.save_page(directory, dict(id='main', window_name='Main', image='images/pages/main.png'))
    assert scanner._reuse_page_identity(directory, image, _page())['id'] == 'main'
    other = _page(identity={'title': 'Other', 'uia_name': 'other'})
    assert scanner._reuse_page_identity(directory, image, other)['id'] != 'main'
    knowledge.save_page(directory, dict(id='main', window_name='main', image='images/pages/main.png'))
    assert scanner._reuse_page_identity(directory, image, _page())['id'] != 'main'


def test_full_scan_does_not_overwrite_other_window_or_reallocate_on_repeat(setup, monkeypatch):
    _layout, native, root = setup
    directory = knowledge.initialize_app('test-process', 'Human display name', root)
    monkeypatch.setattr(knowledge, 'utc_now', lambda: '2026-11-13T00:00:00+00:00')
    monkeypatch.setattr(scanner, '_trim_window_capture', lambda image, bounds, window: (image, bounds))
    monkeypatch.setattr(scanner, 'segment_interface', lambda *a: dict(
        page={'id': 'main', 'name': 'Main'}, regions=[], controls=[]))
    first = scanner.scan_window('Example', '', '', '', 'test', root=root)
    first_image = Path(first['page_image']).read_bytes()
    assert knowledge.read_markdown(directory / 'app.md')[0]['name'] == 'Human display name'
    native['title'] = 'Different dialog'
    window = scanner.uiautomation.ControlFromHandle(1)
    window.Name = 'dialog_Other'
    second = scanner.scan_window('Different dialog', '', '', '', 'test', root=root)
    assert first['page_id'] != second['page_id']
    assert first['page_image'] != second['page_image']
    assert Path(first['page_image']).read_bytes() == first_image
    third = scanner.scan_window('Different dialog', '', '', '', 'test', root=root)
    assert third['page_id'] == second['page_id']
    assert third['scan_id'] != second['scan_id']
    directory = knowledge.app_dir(first['app_id'], root)
    assert len(list(knowledge.iter_records(directory, 'page'))) == 2
    revisions = list((directory / 'history/pages' / second['page_id']).glob('*.md'))
    assert revisions
    previous = knowledge.read_markdown(revisions[0])[0]
    assert previous['annotated_image'] != ('images/pages/'+second['page_id']+'.annotated.png')
    assert (directory / previous['annotated_image']).is_file()


def test_scan_writes_native_title_without_changing_uia_location(setup):
    from test_manual_region_scan import run
    layout, native, root = setup
    native['title'] = 'Native caption'
    window = scanner.uiautomation.ControlFromHandle(1)
    window.Name = 'ProviderObjectName'
    result = run(layout, root)
    index = knowledge.load_index(knowledge.app_dir(result['app_id'], root))
    assert index['pages'][0]['window_name'] == 'Native caption'
    assert index['pages'][0]['window_identity']['uia_name'] == 'ProviderObjectName'
    assert all(r['app_name'] == 'Native caption' for r in index['controls'])
    assert all(r['location']['Xpath'][0]['Name'] == 'Example' for r in index['controls'])


def test_manual_rescan_keeps_previous_screenshot_and_annotation_in_history(setup, monkeypatch):
    from test_manual_region_scan import run
    layout, _native, root = setup
    first = run(layout, root)
    original = Path(first['page_image']).read_bytes()
    annotation = Path(first['annotated_page_image']).read_bytes()
    monkeypatch.setattr(scanner, '_screenshot_window', lambda *a: Image.new('RGB', (400, 300), '#eeeeee'))
    second = run(layout, root)
    assert second['page_id'] == first['page_id']
    assert second['page_image'] != first['page_image']
    assert Path(first['page_image']).read_bytes() == original
    directory = knowledge.app_dir(first['app_id'], root)
    revisions = [knowledge.read_markdown(p)[0] for p in (directory/'history/pages'/first['page_id']).glob('*.md')]
    previous = next(r for r in revisions if r['scan_id'] == first['scan_id'])
    assert (directory / previous['image']).read_bytes() == original
    assert (directory / previous['annotated_image']).read_bytes() == annotation


def test_confirmed_permission_failure_stops_before_activation_or_vault_write(setup, monkeypatch):
    from test_manual_region_scan import run
    layout, _native, root = setup
    monkeypatch.setattr(native_windows, 'check_uia_access', Mock(side_effect=ActionError(
        'operation_blocked', 'integrity mismatch', reason='uia_integrity_mismatch')))
    activate = Mock(side_effect=AssertionError('No activation on denied scan'))
    monkeypatch.setattr(scanner, '_activate_window', activate)
    with pytest.raises(ActionError):
        run(layout, root)
    assert not (knowledge.app_dir('chosen-app', root)/'app.md').exists()
    activate.assert_not_called()


def test_mcp_preflight_uses_legacy_uia_identity_even_with_shadow_native_title(app, monkeypatch, tmp_path):
    window, _button, entries, bind, location = app
    other = FakeControl('Other', 'WindowControl', (0, 0, 400, 300))
    other.ClassName = 'QtWindow'
    entries.insert(0, dict(handle=77, title='SimuNPS', class_name='QtWindow', visible=True))
    bind.side_effect = lambda h: other if h == 77 else window
    monkeypatch.setattr(ui_cli, '_window_rect', lambda *a: scanner._rect(window))
    monkeypatch.setattr(ui_cli, '_screenshot_window', lambda *a: Image.new('RGB', (400, 300)))
    monkeypatch.setattr(ui_cli, '_trim_window_capture', lambda im, rect, win: (im, rect))
    context = Mock(return_value='verified-target')
    monkeypatch.setattr(ui_cli, '_resolve_verified_control_in_context', context)
    record = dict(app_name='SimuNPS', location=deepcopy(location))
    assert ui_cli._resolve_verified_control(tmp_path, record) == 'verified-target'
    assert context.call_args.args[2] is window
    assert record['location'] == location


def test_legacy_visual_record_uses_uia_alias_and_new_record_uses_native_title(app):
    window, _, entries, _, _ = app
    entries[0]['title'] = '目标软件'
    assert scanner._find_window(dict(app_name='SimuNPS', location={'Xpath': [], 'WindowName': 'SimuNPS'})) is window
    assert scanner._find_window(dict(app_name='目标软件', location={'Xpath': [], 'WindowName': '目标软件'})) is window


def test_native_caption_shortlists_qt_window_whose_provider_class_differs(app):
    window, _, entries, bind, location = app
    window.ClassName = 'KCustomDialog'
    location['Xpath'][0]['ClassName'] = 'KCustomDialog'
    entries.insert(0, dict(handle=77, title='Unresponsive other app', class_name='OtherClass', visible=True))
    def selected_only(handle):
        assert handle == window.NativeWindowHandle
        return window
    bind.side_effect = selected_only
    assert scanner._find_window(dict(app_name='SimuNPS-EMT', location=location)) is window
    bind.assert_called_once_with(window.NativeWindowHandle)


@pytest.fixture
def observation(app, monkeypatch):
    window, button, entries, bind, location = app
    entries[0].update(process_id=123, owner_handle=0, z_order=0, foreground=True,
                      bounds=scanner._rect(window))
    monkeypatch.setattr(scanner, '_activate_window', lambda *a: None)
    monkeypatch.setattr(scanner, '_window_rect', lambda win: scanner._rect(win))
    captures = []
    def capture(bounds):
        captures.append(dict(bounds))
        return Image.new('RGB', (bounds['width'], bounds['height']), 'white')
    monkeypatch.setattr(scanner, '_screenshot_window', capture)
    monkeypatch.setattr(scanner, '_trim_window_capture', lambda im, rect, win: (im, rect))
    monkeypatch.setattr(interaction_learning.Controller, 'suspend_highlights', lambda: nullcontext())
    monkeypatch.setattr(interaction_learning, '_capture_virtual_desktop',
                        lambda: (Image.new('RGB', (800, 600)), dict(left=0, top=0)))
    monkeypatch.setattr(native_windows, 'visible_window_bounds', lambda h, rect, **kw: rect)
    return window, entries, bind, location, captures


def test_effect_learning_stays_on_original_window_when_about_dialog_opens(observation, tmp_path, monkeypatch):
    window, entries, bind, location, _captures = observation
    directory = knowledge.initialize_app('example', 'SimuNPS', tmp_path)
    record = {**_record(), 'app_name': 'SimuNPS', 'location': location}
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)
    popup = FakeControl('dialog_AboutWidget', 'WindowControl', (50, 60, 250, 200))
    popup.NativeWindowHandle = 77
    popup.ProcessId = 123
    bound_handles = []
    def bind_control(handle):
        bound_handles.append(handle)
        return popup if handle == 77 else window
    bind.side_effect = bind_control
    def execute(*a, **kw):
        entries[0]['foreground'] = False
        entries.insert(0, dict(handle=77, title='关于 SimuNPS-EMT', class_name='QtWindow',
            process_id=123, owner_handle=window.NativeWindowHandle, visible=True,
            z_order=0, foreground=True, bounds=scanner._rect(popup)))
        return {'ok': True, 'action_executed': True}
    execute_mock = Mock(side_effect=execute)
    monkeypatch.setattr(interaction_learning.ui_cli, 'execute', execute_mock)
    monkeypatch.setattr(interaction_learning, '_popup_controls', lambda *a: [])
    monkeypatch.setattr(interaction_learning, '_interpret_effects', lambda *a: dict(
        effects=[], summary='About opened', confidence=.95, success_condition='About visible'))
    result = interaction_learning.learn_command_effect(directory, record['command']+'.click', '', '', '', 'test')
    assert result['ok'] and result['status'] == 'effect-observed', result
    execute_mock.assert_called_once()
    assert set(bound_handles) == {window.NativeWindowHandle}
    saved = knowledge.list_interactions(directory)[0]
    assert saved['before']['window_name'] == saved['after']['window_name'] == 'SimuNPS-EMT'
    assert 'target_handle' not in saved['after'] and 'target_process_id' not in saved['after']
    assert saved['after']['window_name'] == 'SimuNPS-EMT'
    assert result['window_changes']['added'][0]['handle'] == 77


def test_closed_observation_window_never_rebinds_same_title_or_reads_old_pixels(observation):
    window, entries, bind, _, captures = observation
    before = interaction_learning.capture_snapshot('SimuNPS-EMT')
    entries[0] = {**entries[0], 'handle': 77}
    bind.reset_mock()
    captured = len(captures)
    after = interaction_learning.capture_snapshot(before)
    assert after['target_present'] is False
    assert after['target_handle'] == window.NativeWindowHandle
    assert after['target_image_role'] == 'closed-window-placeholder'
    assert interaction_learning._controls_in_changed_regions(after, after['window_rect'], [{}]) == []
    assert len(captures) == captured
    bind.assert_not_called()


def test_reused_handle_with_other_process_is_rejected(observation):
    _, entries, _, _, _ = observation
    before = interaction_learning.capture_snapshot('SimuNPS-EMT')
    entries[0]['process_id'] = 999
    with pytest.raises(ActionError) as failure:
        interaction_learning.capture_snapshot(before)
    assert failure.value.reason == 'session_window_replaced'


def test_observation_failure_reports_executed_input_and_keeps_attempt_history(observation, monkeypatch, tmp_path):
    _, _, _, location, _ = observation
    directory = knowledge.initialize_app('example', 'SimuNPS', tmp_path)
    record = {**_record(), 'app_name': 'SimuNPS', 'location': location}
    knowledge.save_control(directory, record)
    knowledge.rebuild_index(directory)
    execute = Mock(return_value={'ok': True, 'action_executed': True})
    monkeypatch.setattr(interaction_learning.ui_cli, 'execute', execute)
    monkeypatch.setattr(interaction_learning, 'wait_for_stability', Mock(side_effect=RuntimeError('provider unavailable')))
    result = interaction_learning.learn_command_effect(directory, record['command']+'.click', '', '', '', 'test')
    assert result['ok'] is False and result['action_executed'] is True and result['retryable'] is False
    assert result['code'] == 'result_unconfirmed'
    assert knowledge.list_interactions(directory)[0]['status'] == 'observation-failed'
    execute.assert_called_once()


def test_integrity_preflight_detects_low_to_medium_not_only_elevation(monkeypatch):
    monkeypatch.setattr(native_windows, 'process_integrity_level', lambda pid: 8192 if pid == 42 else 4096)
    monkeypatch.setattr(native_windows, 'process_ui_access', lambda pid: False)
    with pytest.raises(ActionError) as failure:
        native_windows.check_uia_access(42)
    assert failure.value.reason == 'uia_integrity_mismatch'
    assert '4096' in str(failure.value) and '8192' in str(failure.value)


def test_explicit_ui_access_is_not_misdiagnosed_as_insufficient_integrity(monkeypatch):
    monkeypatch.setattr(native_windows, 'process_integrity_level', lambda pid: 12288 if pid == 42 else 8192)
    monkeypatch.setattr(native_windows, 'process_ui_access', lambda pid: True)
    assert native_windows.check_uia_access(42)['ui_access'] is True


def test_unknown_permission_is_not_assumed_to_be_access_denied(monkeypatch):
    monkeypatch.setattr(native_windows, 'process_integrity_level', Mock(side_effect=PermissionError('unavailable')))
    assert native_windows.check_uia_access(42)['status'] == 'unknown'


def test_old_library_without_binding_accepts_explicit_handle_only(monkeypatch):
    entry = dict(handle=42, title='App', class_name='QtWindow', visible=True)
    monkeypatch.setattr(native_windows, 'enumerate_windows', lambda **kw: [entry])
    with pytest.raises(ValueError, match='绑定缺失'):
        library_tools._window({'name': 'Display name'}, 0)
    assert library_tools._window({'name': 'Display name'}, 42) == entry
    with pytest.raises(ValueError, match='无法唯一确定'):
        library_tools._window({'name': 'Display name'}, 99)


def test_dimension_guard_includes_both_sizes_and_handle():
    with pytest.raises(ValueError) as failure:
        manual_layout.segmentation(dict(window_handle=42, image_size={'width': 1920, 'height': 1030}),
                                   {'width': 1918, 'height': 1028})
    assert all(value in str(failure.value) for value in ('1920', '1030', '1918', '1028', '42'))
