from __future__ import annotations

from dataclasses import dataclass


@dataclass
class FakeRect:
    left: int
    top: int
    right: int
    bottom: int

    def width(self) -> int:
        return self.right - self.left

    def height(self) -> int:
        return self.bottom - self.top


class FakeControl:
    def __init__(
        self,
        *,
        name: str = "control",
        class_name: str = "FakeClass",
        control_type: str = "ButtonControl",
        automation_id: str = "",
        rect: FakeRect | None = None,
        parent: FakeControl | None = None,
        enabled: bool = True,
    ) -> None:
        self.Name = name
        self.ClassName = class_name
        self.ControlTypeName = control_type
        self.ControlType = control_type
        self.AutomationId = automation_id
        self.BoundingRectangle = rect or FakeRect(100, 200, 140, 220)
        self.IsEnabled = enabled
        self.NativeWindowHandle = 1
        self.FrameworkId = "Win32"
        self.parent = parent
        self.children: list[FakeControl] = []
        self.events: list[tuple] = []
        if parent is not None:
            parent.children.append(self)

    def Exists(self, *_args, **_kwargs) -> bool:
        return True

    def Refind(self, *_args, **_kwargs) -> bool:
        return True

    def GetParentControl(self):
        return self.parent

    def GetTopLevelControl(self):
        current = self
        while current.parent is not None:
            current = current.parent
        return current

    def GetChildren(self):
        return list(self.children)

    def GetFirstChildControl(self):
        return self.children[0] if self.children else None

    def GetNextSiblingControl(self):
        if self.parent is None:
            return None
        siblings = self.parent.children
        index = siblings.index(self) + 1
        return siblings[index] if index < len(siblings) else None

    def MoveCursorToInnerPos(self, x=None, y=None, ratioX=0.5, ratioY=0.5, **_kwargs):
        rect = self.BoundingRectangle
        relative_x = int(rect.width() * ratioX) if x is None else x
        relative_y = int(rect.height() * ratioY) if y is None else y
        absolute_x = (rect.left if relative_x >= 0 else rect.right) + relative_x
        absolute_y = (rect.top if relative_y >= 0 else rect.bottom) + relative_y
        self.events.append(("move", absolute_x, absolute_y))
        return absolute_x, absolute_y

    def Click(self, x=None, y=None, **kwargs):
        self.events.append(("click", x, y, kwargs))

    def RightClick(self, x=None, y=None, **kwargs):
        self.events.append(("right_click", x, y, kwargs))

    def MiddleClick(self, x=None, y=None, **kwargs):
        self.events.append(("middle_click", x, y, kwargs))

    def DoubleClick(self, x=None, y=None, **kwargs):
        self.events.append(("double_click", x, y, kwargs))

    def SendKeys(self, value):
        self.events.append(("send_keys", value))

