# easy-uiauto
[English](https://github.com/Poggi-Tang/easyautomation/blob/main/README.md) | 简体中文

[![PyPI](https://img.shields.io/pypi/v/easy_uiauto?cacheSeconds=300)](https://pypi.org/project/easy-uiauto/)
[![Python](https://img.shields.io/pypi/pyversions/easy-uiauto?cacheSeconds=300)](https://pypi.org/project/easy-uiauto/)
[![License](https://img.shields.io/github/license/Poggi-Tang/easyautomation)](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE)
[![CI](https://github.com/Poggi-Tang/easyautomation/actions/workflows/ci.yml/badge.svg)](https://github.com/Poggi-Tang/easyautomation/actions/workflows/ci.yml)
[![Publish](https://github.com/Poggi-Tang/easyautomation/actions/workflows/publish.yml/badge.svg)](https://github.com/Poggi-Tang/easyautomation/actions/workflows/publish.yml)

`easy-uiauto` 是一个基于 pyautogui 和 uiautomation 的 UI 自动化测试工具包。

它提供了一套全面的 GUI 自动化 API，包括鼠标控制、键盘输入、窗口管理和控件定位等功能。适用于自动化测试、RPA（机器人流程自动化）以及其他桌面自动化场景。

![logo](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/easy-uiauto.png)

## 功能特点

- 鼠标控制：点击、双击、右键、拖拽
- 键盘输入：文本输入、按键按下/释放、组合键
- 窗口管理：激活、最大化、切换窗口
- 控件定位：基于 XPath 的定位、图像识别
- 视觉反馈：录制时实时高亮显示控件
- 动作录制：录制用户操作并生成脚本
- 富文本字段支持：基于剪贴板的文本输入
- 跨框架支持：Win32、Qt 等多种 UI 框架
- 优先使用 Invoke、Value、Toggle、Selection、Expand/Collapse 等 UIA Pattern
- 提供单一 UIAutomation 工作线程的本地 stdio MCP 服务

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

按需安装开发、桌面集成和 MCP 依赖：

```bash
pip install -e .[dev,integration,mcp]
```

本项目仅支持 Windows，支持 Python 3.11-3.14。

## 命令行

```bash
easy-uiauto --version
easy-uiauto --help
```

## MCP 服务

```bash
pip install -e .[mcp]
easy-uiauto-mcp
```

MCP 服务提供 `list_windows`、`find_control`、`inspect_control` 和
`perform_action`。所有 UIAutomation/COM 对象只在一个专用工作线程中创建和使用，
工具只返回 JSON 快照和短期控件引用。

组合键、鼠标按住/释放和拖拽等高风险动作默认禁用；设置
`EASY_UIAUTO_MCP_ALLOW_HIGH_RISK=1` 后启用。图片路径默认禁用；设置
`EASY_UIAUTO_MCP_ALLOW_IMAGE_PATHS=1` 后启用。

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
- **Semantic Release**：仅能通过手动触发的 release 工作流执行。
- **Trusted Publishing**：通过 GitHub Actions 向 PyPI 发布，无需手动维护 PyPI Token。
- **构建产物**：同时生成 sdist 和 wheel。

## 本地开发

```bash
pip install -e .[dev]
pytest
ruff check .
```

运行 Win32、Qt、Explorer、桌面和任务栏真实集成测试：

```powershell
$env:EASY_UIAUTO_RUN_WINDOWS_TESTS = "1"
pytest tests/test_windows_integration.py -q -s
```

在安全的 SimuNPS 会话可用时运行可逆冒烟测试：

```powershell
$env:EASY_UIAUTO_RUN_SIMUNPS_TESTS = "1"
pytest tests/test_simunps_integration.py -q -s
```

SimuNPS 测试只临时修改模型搜索文本和“消息”过滤开关，并在 `finally` 中恢复；
不会导入、编译、运行、清空或关闭模型。

物理鼠标/键盘回退依赖 Windows 前台输入权限。如果宿主阻止 `SendInput` 或屏幕截图，
语义 UIAutomation Pattern 仍可能正常工作。

## 使用示例

更多示例请参考 `demo/` 目录中的测试文件或查看源代码中的文档字符串。

## 许可证

MIT License，详见 [LICENSE](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE)。

## 联系方式

扫描以下二维码添加我的微信：

![微信二维码](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/or_code.bmp)
