from types import SimpleNamespace

import pytest
from PIL import Image

from easy_uiauto.mcp import control_observation as observation, knowledge, scanner
from easy_uiauto.mcp.results import ActionError


def rect(left=0, top=0, right=100, bottom=80):
    return dict(left=left, top=top, right=right, bottom=bottom, width=right-left, height=bottom-top)


@pytest.fixture
def frames(tmp_path, monkeypatch):
    directory = knowledge.initialize_app("demo", "Demo", tmp_path)
    main = object()
    control = SimpleNamespace(IsOffscreen=False)
    bounds = rect()
    target = rect(10, 10, 40, 30)
    image = Image.new("RGB", (100, 80), "white")
    image.paste("black", (10, 10, 40, 30))
    monkeypatch.setattr(scanner, "_window_rect", lambda _w: bounds)
    monkeypatch.setattr(scanner, "_rect", lambda _c: target)
    monkeypatch.setattr(scanner, "_screenshot_window", lambda _r: image.copy())
    return directory, main, control, bounds, target


def test_target_and_context_use_same_pre_action_pixels_and_are_immutable(frames):
    directory, main, control, _bounds, _target = frames
    before, crop = observation.capture(directory, control, main, phase="before-action")
    assert before["phase"] == "before-action" and crop.getextrema() == ((0, 0),)*3
    with Image.open(directory / before["context_image"]) as context:
        assert context.crop((10, 10, 40, 30)).tobytes() == crop.tobytes()
    after = observation.capture_after(directory, main)
    assert after["phase"] == "after-action" and "normalized_rect" not in after
    # Identical frame bytes share storage instead of growing one file per click.
    assert before["context_image"] == after["context_image"]


def test_disappearing_control_stops_capture_without_saving_background(frames, monkeypatch):
    directory, main, control, _bounds, _target = frames
    def disappearing(_bounds):
        control.IsOffscreen = True
        return Image.new("RGB", (100, 80), "white")
    monkeypatch.setattr(scanner, "_screenshot_window", disappearing)
    with pytest.raises(ActionError, match="发生变化"):
        observation.capture(directory, control, main, phase="before-action")
    assert not list((directory / "images/interactions").glob("*.png"))


def test_result_waits_for_menu_fade_and_never_reads_expired_target(frames, monkeypatch):
    directory, main, _control, _bounds, _target = frames
    clock = {"now": 0.}
    monkeypatch.setattr(observation, "time", SimpleNamespace(
        monotonic=lambda: clock["now"], sleep=lambda seconds: clock.update(now=clock["now"]+seconds)))
    def animated(_bounds):
        return Image.new("RGB", (100, 80), "black" if clock["now"] < .16 else "white")
    monkeypatch.setattr(scanner, "_screenshot_window", animated)
    result = observation.capture_after(directory, main)
    with Image.open(directory / result["context_image"]) as image:
        assert image.getextrema() == ((255, 255),)*3
    assert result["settled"] and result["settle_seconds"] >= .28


def test_popup_context_reframes_main_regions_when_target_is_outside_main():
    main, target = rect(), rect(-20, 0, 10, 20)
    union = observation.union_bounds(main, target)
    description = observation.describe("frame.png", union, target, main, phase="located")
    assert union["width"] == 120
    projected = observation.reframe(dict(left=0, top=0, right=1, bottom=.5), description["main_rect"])
    assert projected["left"] == pytest.approx(20/120, abs=1e-6)
    assert projected["bottom"] == .5
