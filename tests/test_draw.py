from __future__ import annotations

from conftest import FakeControl, FakeRect
from easy_uiauto import draw


def test_visible_rectangle_intersects_all_ancestors():
    window = FakeControl(rect=FakeRect(0, 0, 100, 100))
    panel = FakeControl(rect=FakeRect(10, 10, 90, 90), parent=window)
    control = FakeControl(rect=FakeRect(0, 20, 120, 80), parent=panel)
    assert draw.get_visible_rect_map_by_control(control) == {
        "top": 20,
        "left": 10,
        "bottom": 80,
        "right": 90,
        "width": 80,
        "height": 60,
    }


def test_descendant_at_sample_point_is_not_occlusion(monkeypatch):
    target = FakeControl(rect=FakeRect(0, 0, 100, 100))
    child = FakeControl(rect=FakeRect(10, 10, 90, 90), parent=target)
    monkeypatch.setattr(draw, "get_control_info", lambda _x, _y: ([], child))
    assert draw.get_control_visibl(target) is False


def test_ancestor_at_sample_point_is_not_occlusion(monkeypatch):
    ancestor = FakeControl(rect=FakeRect(0, 0, 100, 100))
    target = FakeControl(rect=FakeRect(10, 10, 90, 90), parent=ancestor)
    monkeypatch.setattr(draw, "get_control_info", lambda _x, _y: ([], ancestor))
    assert draw.get_control_visibl(target) is False


def test_unrelated_control_at_sample_point_is_occlusion(monkeypatch):
    target = FakeControl(rect=FakeRect(0, 0, 100, 100))
    occluder = FakeControl(rect=FakeRect(0, 0, 100, 100))
    monkeypatch.setattr(draw, "get_control_info", lambda _x, _y: ([], occluder))
    assert draw.get_control_visibl(target) is True


def test_destroy_does_not_quit_external_root():
    events = []
    root = type("Root", (), {"quit": lambda self: events.append("quit")})()
    box = draw.ScreenLineBox.__new__(draw.ScreenLineBox)
    box.root = root
    box._owns_root = False
    box._tracking_active = True
    box._destroyed = False
    box._after_id = None
    box.win_top = box.win_right = box.win_bottom = box.win_left = type(
        "Window", (), {"destroy": lambda self: events.append("destroy")}
    )()

    box.destroy()

    assert "quit" not in events


def test_tick_reschedules_after_transient_error(monkeypatch):
    scheduled = []

    class Root:
        def after(self, interval, callback):
            scheduled.append((interval, callback))
            return "after-id"

    box = draw.ScreenLineBox.__new__(draw.ScreenLineBox)
    box.root = Root()
    box.mode = "track"
    box.interval_ms = 25
    box.topmost = False
    box._destroyed = False
    box._after_id = "active"
    box._tracking_active = True
    box._control = object()
    box._rect_snapshot = None
    box._tick_counter = 0
    monkeypatch.setattr(
        draw,
        "get_visible_rect_map_by_control",
        lambda _control: (_ for _ in ()).throw(RuntimeError("transient")),
    )
    monkeypatch.setattr(draw, "push_message", lambda _message: None)

    box._tick()

    assert box._after_id == "after-id"
    assert scheduled == [(25, box._tick)]


def test_rect_snapshot_survives_track_tick(monkeypatch):
    rect = {"left": 1, "top": 2, "right": 11, "bottom": 12, "width": 10, "height": 10}
    updated = []
    box = draw.ScreenLineBox.__new__(draw.ScreenLineBox)
    box.mode = "track"
    box.topmost = False
    box._destroyed = False
    box._after_id = None
    box._tracking_active = True
    box._control = None
    box._rect_snapshot = rect
    box._tick_counter = 0
    box._update_from_rect_map = lambda value, **_kwargs: updated.append(value)
    box._schedule_tick = lambda: None

    box._tick()

    assert updated == [rect]
