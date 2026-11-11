"""A hidden selection must reappear at the next control's actual bounds."""
import os
import subprocess
import sys


def test_manual_highlight_recovers_after_hover_clear_and_follows_parent_child():
    # An isolated interpreter avoids repeated Tk initializations across GUI tests.
    script = r'''
import tkinter as tk
from easy_uiauto.draw import ScreenLineBox

root = tk.Tk(); root.withdraw()
box = ScreenLineBox(root=root, mode="manual", color="#ff1744", line_width=2)
def check(left, top, width, height):
    box.set_rect(dict(left=left, top=top, right=left+width, bottom=top+height,
                      width=width, height=height))
    root.update()
    expected = [(left,top,width,2), (left+width-2,top,2,height),
                (left,top+height-2,width,2), (left,top,2,height)]
    for bar, rect in zip(box._wins(), expected):
        assert bar.winfo_viewable(), "Selected control's outline is hidden"
        assert float(bar.attributes("-alpha")) == 1.0
        assert bar.cget("bg") == "#ff1744"
        actual = (bar.winfo_x(),bar.winfo_y(),bar.winfo_width(),bar.winfo_height())
        assert actual == rect, (actual,rect)
try:
    root.update()
    assert not any(bar.winfo_viewable() for bar in box._wins())
    check(120,120,120,38)
    for index in range(3):
        box.set_visible(False); root.update()
        assert not any(bar.winfo_viewable() for bar in box._wins())
        check(260+index*10,160,120,38)  # Move the mouse to another control.
        check(70,80,720,500)  # Wheel up to its parent.
        check(260+index*10,160,120,38)  # Wheel back down to that child.
    box.set_visible(False); root.update()
    assert not any(bar.winfo_viewable() for bar in box._wins())
finally:
    box.destroy(); root.destroy()
'''
    result = subprocess.run([sys.executable, "-c", script], capture_output=True, text=True,
                            timeout=25, env={**os.environ, "PYTHONIOENCODING":"utf-8"})
    assert result.returncode == 0, result.stdout + result.stderr
