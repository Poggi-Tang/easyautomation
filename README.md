# easy-uiauto

English | [简体中文](https://github.com/Poggi-Tang/easyautomation/blob/main/README.zh-CN.md)

[![PyPI](https://img.shields.io/pypi/v/easy_uiauto?cacheSeconds=300)](https://pypi.org/project/easy-uiauto/)
[![Python](https://img.shields.io/pypi/pyversions/easy-uiauto?cacheSeconds=300)](https://pypi.org/project/easy-uiauto/)
[![License](https://img.shields.io/github/license/Poggi-Tang/easyautomation)](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE)
[![CI](https://github.com/Poggi-Tang/easyautomation/actions/workflows/ci.yml/badge.svg)](https://github.com/Poggi-Tang/easyautomation/actions/workflows/ci.yml)
[![Publish](https://github.com/Poggi-Tang/easyautomation/actions/workflows/publish.yml/badge.svg)](https://github.com/Poggi-Tang/easyautomation/actions/workflows/publish.yml)

`easy-uiauto` is a UI automation toolkit based on pyautogui and uiautomation.

It provides a comprehensive set of APIs for GUI automation, including mouse control, keyboard input, 
window management, and control location. It is suitable for automated testing, RPA (Robotic Process 
Automation), and other desktop automation scenarios.

![logo](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/easy-uiauto.png)

## Features

- Mouse control: click, double-click, right-click, drag and drop
- Keyboard input: text input, key press/release, combination keys
- Window management: activate, maximize, switch windows
- Control location: XPath-based positioning, image recognition
- Visual feedback: real-time control highlighting during recording
- Action recording: record user interactions and generate scripts
- Rich text field support: clipboard-based text input
- Cross-framework support: Win32, Qt, and other UI frameworks
- UI Automation Pattern-first actions for Invoke, Value, Toggle, Selection, and Expand/Collapse
- Local stdio MCP server with a dedicated UIAutomation worker thread

## Installation

Install from PyPI:

```bash
pip install easy-uiauto
```

Or install from source:

```bash
git clone https://github.com/Poggi-Tang/easyautomation.git
cd easyautomation
pip install -e .
```

Install development, desktop-integration, or MCP dependencies as needed:

```bash
pip install -e .[dev,integration,mcp]
```

The package supports Windows and Python 3.11-3.14.

## Command Line

```bash
easy-uiauto --version
easy-uiauto --help
```

## MCP Server

Install the optional dependency and start the local stdio server:

```bash
pip install -e .[mcp]
easy-uiauto-mcp
```

The MCP server exposes discovery and inspection tools (`list_windows`, `find_control`,
`inspect_control`, `list_children`, `get_control_tree`), cache/reference maintenance
(`cache_stats`, `clear_caches`, `invalidate_control`), recording and highlight sessions,
and `perform_action`. UIAutomation/COM objects stay on one worker thread; tools return JSON
snapshots and short-lived control references only. Tk highlight overlays run in a separate
interpreter process so they can be started and stopped safely from stdio MCP clients.

`perform_action` supports dry runs, before/after snapshots, an optional observed control,
and dotted-path postconditions. A successful provider return is not treated as proof that
the desktop state changed: value, toggle, selection, and expand/collapse actions are checked
against the resulting UI Automation state whenever the provider exposes it.

High-risk shortcuts, held mouse buttons, and drag operations are blocked by default.
Set `EASY_UIAUTO_MCP_ALLOW_HIGH_RISK=1` to enable them. Image paths are blocked unless
`EASY_UIAUTO_MCP_ALLOW_IMAGE_PATHS=1` is set. Targets that look like run, compile, delete,
clear, close, or import actions require `confirm_high_impact=true` per call, or
`EASY_UIAUTO_MCP_ALLOW_HIGH_IMPACT=1`. Global input recording is disabled unless the MCP
child process is explicitly started with `EASY_UIAUTO_MCP_ALLOW_RECORDING=1`.

Some Qt accessibility providers report success without committing non-editable combo-box
selection or tree expansion. The controller detects the unchanged state and uses a bounded
window-message fallback against the control's owning Qt window or popup, without moving the
physical cursor; the action is reported successful only after UI Automation exposes the target
state. Traditional physical mouse/keyboard fallback still depends on Windows foreground input
permissions and cannot be guaranteed in a session that blocks `SendInput`.

## Quick Start

### Basic Control Operations

```python
from easy_uiauto.ctrl import Controller

# Left click on a control
Controller.left_click(
    ActionTitle="Click OK Button",
    WindowName="My Application",
    Name="OK",
    ClassName=None,
    ControlType="ButtonControl",
    foundIndex=0,
    AutomationId="",
    Xpath=[],
    Img="",
    PARAMETERS={}
)

# Input text into a field
Controller.input_text(
    ActionTitle="Enter Username",
    WindowName="Login Dialog",
    Name="Username",
    ClassName=None,
    ControlType="EditControl",
    foundIndex=0,
    AutomationId="",
    Xpath=[],
    Img="",
    PARAMETERS={"输入文本": "test_user"}
)

# Keyboard shortcut
Controller.key_group(
    ActionTitle="Save File",
    WindowName="Notepad",
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

### Recording User Actions

```python
from easy_uiauto.record import run_record

# Start recording user actions
run_record(write_file=True)
# Press ESC to stop recording
# Generated script will be saved to Record{timestamp}.py
```

## Project Structure

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
│       ├── ctrl.py          # Core controller (mouse/keyboard actions)
│       ├── draw.py          # Visual feedback (control highlighting)
│       ├── record.py        # Action recording
│       └── utils.py         # Utility functions (control location, caching)
├── tests/
├── CHANGELOG.md
├── LICENSE
├── README.md
├── README.zh-CN.md
└── pyproject.toml
```

## Release Automation

This repository is prepared for a professional Python package workflow:

- **CI** runs lint and tests on push and pull request.
- **Semantic Release** is available only through the manually triggered release workflow.
- **Trusted Publishing** publishes to PyPI from GitHub Actions without a PyPI API token.
- **Build artifacts** include both source distribution and wheel.

## Development

```bash
pip install -e .[dev]
pytest
ruff check .
```

Run deterministic native Win32, Qt, Explorer, desktop/taskbar integration tests:

```powershell
$env:EASY_UIAUTO_RUN_WINDOWS_TESTS = "1"
pytest tests/test_windows_integration.py -q -s
```

Run the reversible SimuNPS smoke test only while a safe SimuNPS session is available:

```powershell
$env:EASY_UIAUTO_RUN_SIMUNPS_TESTS = "1"
pytest tests/test_simunps_integration.py -q -s
```

The SimuNPS test changes only the model-search text and message-filter toggle and restores
both values in a `finally` block. It does not import, compile, run, clear, or close models.

Physical mouse/keyboard fallback depends on Windows foreground-input permissions. Semantic
UIAutomation Patterns can still work when the host blocks `SendInput` or screen capture.

## Usage Examples

For more examples, please refer to the test files in the `demo/` directory or check the docstrings in the source code.

## License

MIT License. See [LICENSE](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE).

## Contact

Scan the QR code to add me on WeChat:

![WeChat QR Code](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/or_code.bmp)
