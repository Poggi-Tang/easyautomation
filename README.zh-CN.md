# easy-uiauto

![logo](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/easy-uiauto.png)

[English](https://github.com/Poggi-Tang/easyautomation/blob/main/README.md) | 简体中文

[![PyPI](https://img.shields.io/pypi/v/easy_uiauto?cacheSeconds=300)](https://pypi.org/project/easy-uiauto/)
[![Python](https://img.shields.io/pypi/pyversions/easy-uiauto?cacheSeconds=300)](https://pypi.org/project/easy-uiauto/)
[![License](https://img.shields.io/github/license/Poggi-Tang/easyautomation)](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE)
[![CI](https://github.com/Poggi-Tang/easyautomation/actions/workflows/ci.yml/badge.svg)](https://github.com/Poggi-Tang/easyautomation/actions/workflows/ci.yml)
[![Publish](https://github.com/Poggi-Tang/easyautomation/actions/workflows/publish.yml/badge.svg)](https://github.com/Poggi-Tang/easyautomation/actions/workflows/publish.yml)

`easy-uiauto` 是一个基于 pyautogui 和 uiautomation 的 UI 自动化测试工具包。

它提供了一套全面的 GUI 自动化 API，包括鼠标控制、键盘输入、窗口管理和控件定位等功能。适用于自动化测试、RPA（机器人流程自动化）以及其他桌面自动化场景。

## 功能特点

- 鼠标控制：点击、双击、右键、拖拽
- 键盘输入：文本输入、按键按下/释放、组合键
- 窗口管理：激活、最大化、切换窗口
- 控件定位：基于 XPath 的定位、图像识别
- 视觉反馈：录制时实时高亮显示控件
- 动作录制：录制用户操作并生成脚本
- 富文本字段支持：基于剪贴板的文本输入
- 跨框架支持：Win32、Qt 等多种 UI 框架

## 安装

从 PyPI 安装：

```bash
pip install easy-uiauto
```

或从源码安装：

```bash
git clone https://github.com/Poggi-Tang/easyautomation.git
cd easyautomation
pip install -e .
```

## 快速示例

### 基本控件操作

```python
from easy_uiauto.ctrl import Controller

# 左键点击控件
Controller.left_click(
    ActionTitle="点击确定按钮",
    WindowName="我的应用",
    Name="确定",
    ClassName=None,
    ControlType="ButtonControl",
    foundIndex=0,
    AutomationId="",
    Xpath=[],
    Img="",
    PARAMETERS={}
)

# 输入文本
Controller.input_text(
    ActionTitle="输入用户名",
    WindowName="登录对话框",
    Name="用户名",
    ClassName=None,
    ControlType="EditControl",
    foundIndex=0,
    AutomationId="",
    Xpath=[],
    Img="",
    PARAMETERS={"输入文本": "test_user"}
)

# 键盘组合键
Controller.key_group(
    ActionTitle="保存文件",
    WindowName="记事本",
    Name="",
    ClassName=None,
    ControlType="",
    foundIndex=0,
    AutomationId="",
    Xpath=[],
    Img="",
    PARAMETERS={"组合键": "ctrl+s"}
)
```

### 录制用户操作

```python
from easy_uiauto.record import run_record

# 开始录制用户操作
run_record(write_file=True)
# 按 ESC 键停止录制
# 生成的脚本将保存到 Record{时间戳}.py
```

## 项目结构

```text
easyautomation
├── .github/
│   └── workflows/
│       ├── ci.yml
│       ├── publish.yml
│       └── release.yml
├── src/
│   └── easy-uiauto/
│       ├── __init__.py
│       ├── ctrl.py          # 核心控制器（鼠标/键盘操作）
│       ├── draw.py          # 视觉反馈（控件高亮）
│       ├── record.py        # 动作录制
│       └── utils.py         # 工具函数（控件定位、缓存）
├── tests/
├── CHANGELOG.md
├── LICENSE
├── README.md
├── README_CN.md
└── pyproject.toml
```

## 发布自动化

当前仓库已经按照较完整的 Python 开源库流程整理：

- **CI**：在 push 和 pull request 时自动执行 lint 与测试。
- **Semantic Release**：自动更新版本号、CHANGELOG、Tag 和 GitHub Release。
- **Trusted Publishing**：通过 GitHub Actions 向 PyPI 发布，无需手动维护 PyPI Token。
- **构建产物**：同时生成 sdist 和 wheel。

## 本地开发

```bash
pip install -e .[dev]
pytest
ruff check .
```

## 使用示例

更多示例请参考 `demo/` 目录中的测试文件或查看源代码中的文档字符串。

## 许可证

MIT License，详见 [LICENSE](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE)。

## 联系方式

扫描以下二维码添加我的微信：

![微信二维码](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/or_code.bmp)