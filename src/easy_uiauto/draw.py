# -*- coding: utf-8 -*-
# @Name:      draw.py
# @Author:    tang
# @Date:      2025/10/9-18:40
# @depict:
import tkinter as tk
from dataclasses import dataclass
import uiautomation

from utils import get_control_info


# ================= 可见区域算法（输出 rect_map） =================
@dataclass
class RectStruct:
    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self):
        return max(0, self.right - self.left)

    @property
    def height(self):
        return max(0, self.bottom - self.top)

    def is_valid(self):
        return self.width > 0 and self.height > 0

    def to_rect_map(self):
        return {
            'top': self.top,
            'left': self.left,
            'bottom': self.bottom,
            'right': self.right,
            'width': self.width,
            'height': self.height
        }


def _to_rect(rect_like):
    if isinstance(rect_like, (tuple, list)) and len(rect_like) == 4:
        return RectStruct(*rect_like)
    if isinstance(rect_like, RectStruct):
        return rect_like
    if hasattr(rect_like, "left") and hasattr(rect_like, "top") \
            and hasattr(rect_like, "right") and hasattr(rect_like, "bottom"):
        return RectStruct(rect_like.left, rect_like.top, rect_like.right, rect_like.bottom)
    raise TypeError(f"不支持的 rect 类型: {type(rect_like)} -> {rect_like}")


def _intersect_rect(parent_rect, control_rect):
    p = _to_rect(parent_rect)
    c = _to_rect(control_rect)
    r = RectStruct(
        left=max(c.left, p.left),
        top=max(c.top, p.top),
        right=min(c.right, p.right),
        bottom=min(c.bottom, p.bottom),
    )
    if not r.is_valid():
        return RectStruct(0, 0, 0, 0)
    return r


# @timeit
def get_visible_rect_map_by_control(control):
    try:
        c_rect = control.BoundingRectangle
    except Exception:
        c_rect = None

    if c_rect is None:
        return {'top': 0, 'left': 0, 'bottom': 0, 'right': 0, 'width': 0, 'height': 0}

    try:
        parent = control.GetParentControl()
        p_rect = parent.BoundingRectangle if parent else None
    except Exception:
        parent, p_rect = None, None

    if p_rect is None:
        return _to_rect(c_rect).to_rect_map()

    return _intersect_rect(p_rect, c_rect).to_rect_map()


# ================= 绘制类：ScreenLineBox =================
class ScreenLineBox:
    """
    四条“线”= 四个置顶、无边框的 Toplevel 窗口，框出控件的可见区域。
    - 默认：红色、线宽2、alpha=1.0
    - 初始化修复：创建时窗口放屏幕外并设为全透明；首次获得有效矩形后再显示（无闪烁/无左上角小块）。
    - mode="track": 每 interval_ms 跟踪；丢失/遮挡切色但不中断循环。
    """

    LOST_COLOR = "#FFD400"      # 控件丢失时：黄色
    OCCLUDED_COLOR = "#1E90FF"  # ★修改：遮挡时：蓝色
    NORMAL_COLOR = "#FF0000"    # 默认红色
    _OFFSCREEN_GEOM = "1x1+-20000+-20000"

    def __init__(
            self,
            root=None,
            control=None,
            *,
            mode="track",
            interval_ms=30,
            line_width=2,
            color="#FF0000",
            alpha=1.0,
            topmost=True,
            show_time=1000,
            occlusion_check_interval_ms=200,
    ):
        self._destroyed = None
        self._owns_root = False
        if root is None:
            root = tk.Tk()
            root.withdraw()
            self._owns_root = True

        self.root = root
        self.root.update_idletasks()

        self.mode = (mode or "track").lower()
        # 降低默认 interval_ms（更低延迟），但记录线程也可以覆盖此值
        self.interval_ms = max(10, int(interval_ms))
        self.line_w = max(1, int(line_width))
        self.color = color
        self.alpha = float(alpha)
        self.topmost = topmost

        # 遮挡检测频率控制（以 ms 为单位），避免每帧都做昂贵检测
        self.occlusion_check_interval_ms = max(0, int(occlusion_check_interval_ms))
        # 计算每多少 tick 执行一次遮挡检测（至少1次）
        self._occlusion_tick_mod = max(1, (self.occlusion_check_interval_ms // self.interval_ms) or 1)
        self._tick_counter = 0

        # 状态
        self._control = control
        self._tracking_active = False
        self._paused_for_missing = False  # ★保留字段但不再用于暂停
        self._normal_color_cache = self.color
        self._ever_drawn_valid = False
        self._is_shown = False
        self._current_mode = "normal"  # normal / lost / occluded  ★修改：增加模式状态

        # 四条“线”
        self.win_top = self._make_bar()
        self.win_right = self._make_bar()
        self.win_bottom = self._make_bar()
        self.win_left = self._make_bar()

        # 初次绘制
        self._draw_once_from_control()

        # 启动模式
        if self.mode == "once":
            # 仅演示一次，按需显示
            self.root.after(show_time, self.destroy)
        elif self.mode == "track":
            self._tracking_active = True
            self._tick()
        else:
            self.root.after(1000, self.destroy)

    # ---------- 对外 API ----------
    def set_control(self, control):
        self._control = control
        rect_map = get_visible_rect_map_by_control(control) if control is not None else None
        if rect_map and rect_map['width'] > 0 and rect_map['height'] > 0:
            self._enter_normal_mode()  # ★可见先恢复正常色，再根据遮挡判定切蓝
            # 立即更新位置，但延后/降低遮挡检测频率以提升响应
            self._update_from_rect_map(rect_map, run_occlusion_check=False)
            if self.mode == "track":
                self._tracking_active = True
        else:
            self._enter_lost_mode(initial=not self._ever_drawn_valid)

    def set_style(self, *, color=None, line_width=None, alpha=None):
        if color is not None:
            self.color = color
            self._normal_color_cache = color
            self._apply_color_to_all(self._resolve_color_by_mode(self._current_mode))
        if alpha is not None:
            self.alpha = float(alpha)
            for w in self._wins():
                try:
                    w.wm_attributes("-alpha", self.alpha if self._is_shown else 0.0)
                except tk.TclError:
                    pass
        if line_width is not None:
            self.line_w = max(1, int(line_width))
            rect_map = get_visible_rect_map_by_control(self._control) if self._control is not None else None
            if rect_map and rect_map['width'] > 0 and rect_map['height'] > 0:
                self._update_from_rect_map(rect_map, run_occlusion_check=False)

    def pause_tracking(self):
        if self.mode == "track":
            self._tracking_active = False

    def resume_tracking(self):
        if self.mode == "track":
            self._tracking_active = True
            self._enter_normal_mode()
            self._tick()

    def destroy(self):
        self._tracking_active = False
        self._destroyed = True

        if self._owns_root:
            self._owns_root = False

        for w in self._wins():
            try:
                w.destroy()
            except Exception:
                pass
        self.root.quit()

    # ---------- 内部 ----------
    def _wins(self):
        return (self.win_top, self.win_right, self.win_bottom, self.win_left)

    def _make_bar(self):
        w = tk.Toplevel(self.root)
        w.overrideredirect(True)
        w.configure(bg=self.NORMAL_COLOR)
        if self.topmost:
            w.wm_attributes("-topmost", True)

        try:
            w.geometry(self._OFFSCREEN_GEOM)
        except Exception:
            pass
        try:
            w.wm_attributes("-alpha", 0.0)
        except tk.TclError:
            pass
        try:
            w.withdraw()
        except Exception:
            pass
        try:
            w.wm_attributes("-toolwindow", True)
        except tk.TclError:
            pass
        return w

    def _reveal(self):
        for w in self._wins():
            try:
                w.deiconify()
            except Exception:
                pass
            try:
                w.wm_attributes("-alpha", self.alpha)
            except tk.TclError:
                pass
        self._is_shown = True

    def _apply_color_to_all(self, color):
        for w in self._wins():
            w.configure(bg=color)

    def _resolve_color_by_mode(self, mode):
        if mode == "lost":
            return self.LOST_COLOR
        if mode == "occluded":
            return self.OCCLUDED_COLOR
        return self._normal_color_cache  # normal

    def _enter_lost_mode(self, initial=False):
        """
        丢失：仅改成黄框，不暂停更新（满足规则1）
        initial=True 且从未成功绘制过：保持隐藏，避免初始闪烁
        """
        self._current_mode = "lost"  # ★
        self._paused_for_missing = True  # 仅作标记，不再控制循环
        if initial and not self._ever_drawn_valid:
            return
        else:
            self._apply_color_to_all(self._resolve_color_by_mode("lost"))
            if not self._is_shown and self._ever_drawn_valid:
                self._reveal()

    def _enter_occluded_mode(self):
        """遮挡：蓝框（满足规则2）"""
        if self._current_mode != "occluded":
            self._current_mode = "occluded"
            self._apply_color_to_all(self._resolve_color_by_mode("occluded"))

    def _enter_normal_mode(self):
        """正常：红框（满足规则3）"""
        if self._current_mode != "normal":
            self._current_mode = "normal"
            self._apply_color_to_all(self._resolve_color_by_mode("normal"))

    def _draw_once_from_control(self):
        rect_map = get_visible_rect_map_by_control(self._control) if self._control is not None else None
        if not rect_map or rect_map['width'] <= 0 or rect_map['height'] <= 0:
            self._enter_lost_mode(initial=True)
            return
        self._update_from_rect_map(rect_map, run_occlusion_check=True)

    def _place(self, win, x, y, width, height):
        win.geometry(f"{max(1, int(width))}x{max(1, int(height))}+{int(x)}+{int(y)}")

    def _update_from_rect_map(self, rect_map, *, run_occlusion_check: bool):
        """
        更新位置；当空间可见时才进行遮挡检测（满足规则4）
        """
        x1, y1 = rect_map['left'], rect_map['top']
        x2, y2 = rect_map['right'], rect_map['bottom']
        x1, x2 = sorted((int(x1), int(x2)))
        y1, y2 = sorted((int(y1), int(y2)))
        t = self.line_w
        # 顶边
        self._place(self.win_top, x1, y1, max(1, x2 - x1), t)
        # 底边
        self._place(self.win_bottom, x1, max(y1, y2 - t), max(1, x2 - x1), t)
        # 左边
        self._place(self.win_left, x1, y1, t, max(1, y2 - y1))
        # 右边
        self._place(self.win_right, max(x1, x2 - t), y1, t, max(1, y2 - y1))

        if not self._ever_drawn_valid:
            self._ever_drawn_valid = True
            self._enter_normal_mode()
            self._reveal()

        # ★只有在需要并且达到检测周期时才检测遮挡
        do_occlusion = False
        if run_occlusion_check and rect_map['width'] > 0 and rect_map['height'] > 0 and self._control is not None:
            # 根据计数决定是否真正运行昂贵的遮挡检测
            if (self._tick_counter % self._occlusion_tick_mod) == 0:
                do_occlusion = True

        if do_occlusion:
            is_occluded = get_control_visibl(self._control, rect_map)  # True=被遮挡
            if is_occluded:
                self._enter_occluded_mode()
            else:
                self._enter_normal_mode()

    def _tick(self):
        if getattr(self, "_destroyed", False):
            return
        if self.mode != "track":
            return
        if self.topmost:
            for w in self._wins():
                w.wm_attributes("-topmost", True)
                w.lift()

        if self._tracking_active:
            rect_map = get_visible_rect_map_by_control(self._control) if self._control is not None else None
            if rect_map and rect_map['width'] > 0 and rect_map['height'] > 0:
                # 先更新位置，再根据遮挡切色（但实际遮挡检测由周期控制）
                self._tick_counter += 1
                # 仅在周期到点时传入 True，让 _update_from_rect_map 决定是否执行遮挡检测
                self._update_from_rect_map(rect_map, run_occlusion_check=True)
            else:
                # 丢失：变黄，但不暂停；继续 tick
                self._enter_lost_mode(initial=False)
        self.root.after(self.interval_ms, self._tick)

    def run(self):
        if self._owns_root:
            self.root.mainloop()


def check_visibl(control_point, top_window_point, check_map=None):
    if check_map is None:
        check_map = {'left': True, 'right': True, 'top': True, 'bottom': True}
    if check_map.get('left') and control_point['left'] > top_window_point['left']:
        # print(f"左侧被遮挡,控件坐标{control_point},遮挡控件坐标{top_window_point}")
        check_map['left'] = False
    if check_map.get('right') and control_point['right'] < top_window_point['right']:
#         print(f"右侧被遮挡,控件坐标{control_point},遮挡控件坐标{top_window_point}")
        check_map['right'] = False
    if check_map.get('top') and control_point['top'] > top_window_point['top']:
#         print(f"顶部被遮挡,控件坐标{control_point},遮挡控件坐标{top_window_point}")
        check_map['top'] = False
    if check_map.get('bottom') and control_point['bottom'] < top_window_point['bottom']:
#         print(f"底部被遮挡,控件坐标{control_point},遮挡控件坐标{top_window_point}")
        check_map['bottom'] = False
    return check_map


def get_pos_top_window_point(x, y, control):
    _, pos_control = get_control_info(int(x), int(y))
    try:
        if pos_control.BoundingRectangle == control.BoundingRectangle:
            return False
        top_window = pos_control.GetTopLevelControl()
        point_top = get_visible_rect_map_by_control(top_window)
        return point_top
    except Exception as e:
#         print(e)
        return False


# @timeit
def get_control_visibl(control, rect_map=None):
    """
    返回 True 表示被遮挡；False 表示未遮挡。
    仅当传入 rect_map 为有效可见矩形时，才有意义。
    """
    if rect_map is None:
        point = get_visible_rect_map_by_control(control)
    else:
        point = rect_map

    if not point or point.get('width', 0) <= 0 or point.get('height', 0) <= 0:
        return False  # 空间上不可见，这里不做遮挡判断

    topx, topy = point['width'] / 2 + point['left'], point['top'] + 7
    leftx, lefty = point['left'] + 7, point['height'] / 2 + point['top']
    rightx, righty = point['right'] - 7, point['height'] / 2 + point['top']
    bottomx, bottomy = point['width'] / 2 + point['left'], point['bottom'] - 7

    check_map = {'left': True, 'right': True, 'top': True, 'bottom': True}

    point_top = get_pos_top_window_point(topx, topy, control)
    if point_top:
        check_map = check_visibl(point, point_top, check_map)

    point_left = get_pos_top_window_point(leftx, lefty, control)
    if point_left:
        check_map = check_visibl(point, point_left, check_map)

    point_right = get_pos_top_window_point(rightx, righty, control)
    if point_right:
        check_map = check_visibl(point, point_right, check_map)

    point_bottom = get_pos_top_window_point(bottomx, bottomy, control)
    if point_bottom:
        check_map = check_visibl(point, point_bottom, check_map)

    # 任意一侧为 False 即视为被遮挡
    return not all(check_map.values())


# ================== 示例用法 ==================
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    ctrl2 = uiautomation.WindowControl(Name='SimuLab5.1')  # 按你的场景
    box_track = ScreenLineBox(root, ctrl2, mode="track", interval_ms=100)

    root.mainloop()
