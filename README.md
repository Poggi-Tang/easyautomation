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

### MCP Server

Install the optional MCP dependencies when using easy-uiauto from an MCP client:

```bash
pip install "easy-uiauto[mcp]"
```

Install local OCR and image-template fallback support when needed:

```bash
pip install "easy-uiauto[mcp,vision]"
```

The `vision` extra provides OpenCV template matching and the Python Tesseract
adapter. OCR also requires the system Tesseract executable and relevant language
data (for example `eng` or `chi_sim`). The MCP tools are
`find_control_by_image`, `click_by_image`, `find_text_on_screen`, and
`click_text_on_screen`.

Remote multimodal location does not run a local model or require an AI SDK. Set
an OpenAI-compatible vision endpoint and credentials, then use
`find_control_by_vision` or `click_by_vision`:

```bash
EASY_UIAUTO_VISION_API_URL=https://your-api.example/v1/chat/completions
EASY_UIAUTO_VISION_API_KEY=your-api-key
EASY_UIAUTO_VISION_MODEL=your-vision-model
```

Those tools upload the current screenshot to the configured endpoint only for
that request. Use them as a final fallback after UIA, OCR, or image matching.

The MCP server is part of the library and reuses the same automation APIs:

```bash
easy_uiauto --help
easy_uiauto --version
easy_uiauto
```

Use the following commands to register, inspect, or remove the global MCP
configuration for an installed client. They use the client's own CLI and never
overwrite an existing entry with the same name.

```bash
easy_uiauto --install-codex
easy_uiauto --show-codex-config
easy_uiauto --uninstall-codex

easy_uiauto --install-claude-code
easy_uiauto --show-claude-code-config
easy_uiauto --uninstall-claude-code
```

Codex registration uses its global `config.toml`. Claude Code registration uses
the `user` scope, so it is available to every local project. Restart the client
after installing or removing the server.

The long-running TCP service is also available:

```bash
easy_uiauto_service --help
python -m easy_uiauto.mcp.service --port 9876
```

For MCP client configuration, start the server with `python -m easy_uiauto.mcp.server`.
Control-vector persistence is optional. To enable it, set
`EASY_UIAUTO_CONTROL_VECTOR_DB_DIR` to a directory containing
`control_vector_store.py`; otherwise capture tools still return records but do not persist them.

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
