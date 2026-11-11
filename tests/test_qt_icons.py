"""Exercise Qt's real icon dispatch, DPI scaling, and selection contrast."""
import pytest

pytest.importorskip("PySide6")
from PySide6.QtCore import QSize
from PySide6.QtGui import QColor, QIcon, QPalette
from PySide6.QtWidgets import QApplication

from easy_uiauto.qt_ui.icons import icon


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


@pytest.mark.parametrize("scale", [1, 1.5, 2])
def test_icon_preserves_logical_size_and_renders_at_device_resolution(app, scale):
    pixmap = icon("model-chip").pixmap(QSize(18, 18), scale)
    assert pixmap.size() == QSize(round(18*scale), round(18*scale))
    assert pixmap.deviceIndependentSize().width() == 18
    image = pixmap.toImage()
    assert any(image.pixelColor(x, y).alpha() for x in range(image.width()) for y in range(image.height()))


@pytest.mark.parametrize("background,expected", [("#1E293B", "#ffffff"), ("#EFF6FF", "#2563eb")])
def test_selected_icon_contrasts_with_the_actual_palette(app, background, expected):
    original = app.palette()
    try:
        palette = QPalette(original)
        palette.setColor(QPalette.ColorRole.Highlight, QColor(background))
        app.setPalette(palette)
        selected = icon("add").pixmap(QSize(24, 24), QIcon.Mode.Selected).toImage()
        assert selected.pixelColor(12, 12).name() == expected
    finally:
        app.setPalette(original)


@pytest.mark.parametrize("name", ["add", "stop"])
def test_disabled_color_applies_to_both_strokes_and_fills(app, name):
    image = icon(name).pixmap(QSize(24, 24), QIcon.Mode.Disabled).toImage()
    assert image.pixelColor(12, 12).name() == "#9ca3af"
