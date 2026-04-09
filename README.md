# easy-uiauto

(https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/easy-uiauto.png)

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
- **Semantic Release** updates the version, changelog, tag, and GitHub Release.
- **Trusted Publishing** publishes to PyPI from GitHub Actions without a PyPI API token.
- **Build artifacts** include both source distribution and wheel.

## Development

```bash
pip install -e .[dev]
pytest
ruff check .
```

## Usage Examples

For more examples, please refer to the test files in the `demo/` directory or check the docstrings in the source code.

## License

MIT License. See [LICENSE](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE).

## Contact

Scan the QR code to add me on WeChat:

![WeChat QR Code](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/or_code.bmp)
