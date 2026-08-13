# easy-uiauto

English | [简体中文](https://github.com/Poggi-Tang/easyautomation/blob/main/README.zh-CN.md)

[![PyPI](https://img.shields.io/pypi/v/easy_uiauto?cacheSeconds=300)](https://pypi.org/project/easy-uiauto/)
[![Python](https://img.shields.io/pypi/pyversions/easy-uiauto?cacheSeconds=300)](https://pypi.org/project/easy-uiauto/)
[![License](https://img.shields.io/github/license/Poggi-Tang/easyautomation)](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE)
[![CI](https://github.com/Poggi-Tang/easyautomation/actions/workflows/ci.yml/badge.svg)](https://github.com/Poggi-Tang/easyautomation/actions/workflows/ci.yml)
[![Publish](https://github.com/Poggi-Tang/easyautomation/actions/workflows/publish.yml/badge.svg)](https://github.com/Poggi-Tang/easyautomation/actions/workflows/publish.yml)

![logo](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/easy-uiauto.png)

**Turn a Windows desktop interface into inspectable, reusable UI commands.**

`easy-uiauto` combines a Python automation library, an MCP server, action
recording, and an Obsidian-compatible UI knowledge vault. It can locate controls
through Windows UI Automation, learn the useful parts of a visible application,
and expose verified operations as commands such as
`main.keypad.enter-digit-6.click`.

The project is useful for Windows test automation, RPA, agent tooling, and
building reusable operation knowledge for desktop software. It is Windows-only
and currently released as a `0.x` project.

> ⚠️ Visual scanning and workflow exploration are still evolving. Review learned
> commands before using them for messages, payments, deletion, uploads, account
> changes, or other operations with external effects.

## ✨ What It Includes

- 🖱️ **Desktop automation**: click, double-click, right-click, drag, scroll,
  keyboard input, hotkeys, window activation, and text entry.
- 🎬 **Action recording**: record real operations, highlight the selected control,
  and generate reusable Python actions with a complete `LOCATION`.
- 🔌 **MCP server**: use the same automation surface from Codex, Claude Code, or
  another MCP client.
- 🔎 **Visual-first learning**: identify the important regions and controls in one
  screenshot, then map them back to stable local UIA controls without walking the
  whole tree by default.
- 🧠 **State-aware semantics**: distinguish stable functions from dynamic values,
  retain names and avatars, connect related controls, and explain visible states such
  as a Send button disabled while its message input is empty.
- 🧭 **Semantic UI CLI**: turn verified controls into searchable application
  commands instead of rediscovering coordinates for every task.
- 🖼️ **Control overlays**: number every learned or executable control in one
  click-through overlay, save annotated page images, and preview operation targets.
- 🧪 **Effect learning**: compare before/after screenshots, inspect changed UIA
  properties and new windows, and store an evidence-backed success condition.
- 🗂️ **Readable knowledge**: keep pages, controls, images, interactions, and
  quarantine records as Markdown/YAML/PNG files that work directly with Obsidian
  and version control.
- 📐 **Stable records**: persist structured `LOCATION` identities and normalized
  visual hints, never process IDs, window handles, or absolute desktop rectangles.
- 🛟 **Layered lookup**: unique `AutomationId`, contextual XPath, saved image states,
  local OCR, then opt-in remote vision as the final fallback.
- ⚡ **Fast anchors**: use a unique window-scoped `AutomationId` before XPath, while
  keeping hierarchy as the fallback for missing or duplicate IDs.

## 🚦 Capability Status

| Area | Status | Notes |
| --- | --- | --- |
| Python mouse, keyboard, window, and UIA APIs | Stable | The original library surface |
| Recording and replay | Stable | Produces reusable structured actions |
| MCP tools and client setup | Beta | Codex and Claude Code installers included |
| Visual-first scan and generated UI CLI | Beta | Requires an OpenAI-compatible vision endpoint |
| Before/after effect learning | Beta | Stores local evidence and learned success conditions |
| Recursive workflow exploration | Experimental | Restricted to reversible commands and bounded depth |

Automated workflow exploration intentionally excludes scrolling and dragging.
The base automation library still provides those operations for explicit scripts
and direct MCP calls.

## 📦 Installation

Requirements: Windows and Python 3.11 or newer.

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

## 🔌 MCP Setup

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

For a coarse UIA canvas whose internal elements have no UIA nodes, pass the
control result or its bounds directly to the local rectangle detector:

```text
get_control_at_position(x, y)
detect_visual_elements(rect=<the returned control object>)
```

`detect_visual_elements` captures only the supplied rectangle and returns numbered
class-agnostic boxes, absolute screen rectangles, input-relative normalized rectangles,
containment relationships, confidence, timing, and an annotated preview path. It does
not traverse UIA, run OCR, infer semantics, download a model, or call a remote API.

Remote visual location does not run a local model or require a vendor SDK. Set
an OpenAI-compatible vision endpoint and credentials, then use
`find_control_by_vision` or `click_by_vision`:

```bash
EASY_UIAUTO_VISION_API_URL=https://your-api.example/v1/chat/completions
EASY_UIAUTO_VISION_API_KEY=your-api-key
EASY_UIAUTO_VISION_MODEL=your-vision-model
```

Those tools upload the current screenshot to the configured endpoint only for
that request. Use them as a final fallback after UIA, OCR, or image matching.

The MCP server ships with the library and reuses the same automation APIs:

```bash
easy_uiauto --help
easy_uiauto --version
easy_uiauto
```

The setup commands register, inspect, or remove the MCP entry through the
client's own CLI. Existing entries are not overwritten by the basic installer.

For a complete Codex installation:

```bash
pip install --upgrade "easy-uiauto[mcp,vision]"
easy_uiauto --full-setup-codex \
  --vision-url https://your-api.example/v1/chat/completions \
  --vision-model your-vision-model
```

Full setup asks for a missing API key through hidden input, installs the Python
vision dependencies and Tesseract when needed, replaces only the `easy_uiauto`
Codex entry, installs the two bundled skills, and checks UIA, OCR, and the remote
vision endpoint. Validation uses generated images rather than the current desktop.
Vision settings are read dynamically by a current-version MCP process. Restart
Codex once after installing or updating the package, MCP entry, or skills; changing
only the vision settings does not require a second restart.

For a minimal Codex deployment with remote AI vision, use the quick setup
command. When the API key is not already present in the Windows user
environment, it prompts once through hidden terminal input or a password dialog
for non-interactive agents. It persists the three vision variables, replaces
only the `easy_uiauto` Codex MCP entry, and skips OCR installation and UI tests:

```bash
easy_uiauto --quick-setup-codex \
  --vision-url https://your-api.example/v1/chat/completions \
  --vision-model your-vision-model
```

You can also give Codex this deployment instruction:

> Install the latest `easy-uiauto[mcp,vision]` from PyPI, then run
> `easy_uiauto --full-setup-codex --vision-url URL --vision-model MODEL`; do not
> inspect unrelated projects or search the web; report every validation result
> and elapsed time, then ask me to restart Codex once. Do not request another
> restart only because the vision environment variables changed.

```bash
easy_uiauto --install-codex
easy_uiauto --show-codex-config
easy_uiauto --uninstall-codex

easy_uiauto --install-claude-code
easy_uiauto --show-claude-code-config
easy_uiauto --uninstall-claude-code
```

Codex registration uses its global `config.toml`. Claude Code uses the `user`
scope. Restart the client after installing, updating, or removing the server.

The long-running TCP service is also available:

```bash
easy_uiauto_service --help
python -m easy_uiauto.mcp.service --port 9876
```

## 🧭 Learn an Application

The package includes two Codex skills:

- `$easy-uiauto-learning` scans and explores visible Windows applications.
- `$easy-uiauto-operate` runs commands from an existing application vault.

After setup and one Codex restart, a natural-language request is enough:

```text
Use $easy-uiauto-learning to learn the currently open WeChat window and build
its control knowledge and UI commands.
```

### 🔎 Light Scan

A light scan observes one visible page. It does not click the application.

```bash
easy_uiauto_ui scan "Window title"
easy_uiauto_ui apps
easy_uiauto_ui commands <app-id>
easy_uiauto_ui show <app-id>
```

The default `visual-first` strategy sends one target-window screenshot to the
configured vision endpoint, identifies useful regions and controls, and maps
those coordinates back to local UIA controls. It stores stable `LOCATION`
records, control crops, meaning, aliases, risk, and generated commands. Secondary
monitors are supported. The scan also saves `<page-id>.annotated.png` and briefly
draws the same numbered controls over the live window. Use `--no-overlay` when
visual feedback is not wanted.

MCP scans report their task ID, stage, elapsed time, and control-verification count
every five seconds. If a client stops waiting, call `get_ui_learning_status` before
retrying. A `running` task continues in the MCP process and must not be started twice.
Plain learning requests stop after this scan. Interactive `explore_ui_workflows`
runs only when deep or autonomous exploration is explicitly requested.

Use `--strategy full-uia` for diagnostics when visual targeting is not suitable.
This mode walks the visible UIA tree and performs a separate semantic pass, so it
is slower and produces more records.

### 🧪 Deep Exploration

Deep exploration starts with a light scan, executes known reversible commands,
and records what changed. New pages and dialogs can be scanned recursively.

```text
Use $easy-uiauto-learning to deeply learn the currently open WeChat window with
the safe policy, at most 30 actions, and recursion depth 3. Do not send messages,
delete data, log out, or run any confirmation-required operation.
```

```bash
easy_uiauto_ui explore <app-id> --policy safe --max-actions 30 --max-depth 3
easy_uiauto_ui interactions <app-id>
```

For each operation, the explorer saves before/after screenshots, local pixel
differences, changed UIA properties, new or transient windows, the observed
success condition, and the recovery result. It deduplicates page/command states
and stops when recovery fails or another user/window interrupts the run.

`safe` executes only commands classified as safe. `supervised` may also execute
reversible state-changing commands. External, destructive, and
confirmation-required commands are never explored automatically. Scrolling and
dragging are not implemented in the explorer.

### ▶️ Run Learned Commands

```bash
easy_uiauto_ui search <app-id> "search terms"
easy_uiauto_ui show <app-id> --include executable
easy_uiauto_ui run <app-id> <page.region.control.action> --text "optional text"
easy_uiauto_ui batch <app-id> '["main.keypad.enter-digit-6.click", "main.keypad.enter-addition-operation.click"]'
easy_uiauto_ui learn-effect <app-id> <command> --recover
easy_uiauto_ui teach <app-id> <control-id> "Meaning" intent "Description"
```

Runtime lookup follows unique `AutomationId` → contextual XPath → saved image states → local OCR → opt-in
remote vision. Enable the last step with `--allow-vision-fallback` or
`allow_vision_fallback=true`. Stale, missing, or ambiguous controls are moved to
quarantine instead of being clicked. Single and batch commands preview their
resolved targets in red by default; pass `--no-highlight` or `highlight=false`
to remove the roughly 100 ms visual-preview wait.

`set-text` first verifies UIA value assignment, then falls back to clipboard paste
for Unicode text. It reports success only after an accessible value or exact clipboard
readback confirms the write; a local image change is diagnostic evidence only. The user's
previous text clipboard value is restored after pasting.

Use a batch only for commands on the same stable page. Split the sequence after
navigation. Commands with external or destructive effects require `--confirm`
or `confirm=true`.

### 🗂️ Knowledge Vault

The default vault is `~/easy_uiauto_vault`. Set
`EASY_UIAUTO_KNOWLEDGE_DIR` to use another location. Markdown/YAML and PNG files
are the source of truth:

```text
applications/<app-id>/
├── pages/                 page records
├── regions/               functional regions
├── controls/              verified and observed controls
├── interactions/          before/after effect records
├── images/                page, control, and interaction images
├── quarantine/            stale or ambiguous controls
└── operations/UI-CLI.md   generated command catalog
```

`.easy_uiauto/index.json` is a disposable search cache. Rebuild it with
`easy_uiauto_ui reindex <app-id>` after editing Markdown by hand.

Control and region geometry is stored only as a window-relative `normalized_rect`
with `geometry_role: visual-hint-only`. It is useful for annotations and visual
comparison, but never identifies or clicks a control. Live overlays and operations
resolve the current `LOCATION` first, then use a unique image template or local OCR
fallback. Version 0.6 automatically removes legacy absolute rectangles, PIDs, and
window handles when an older vault index is rebuilt.

### 🔐 Privacy and Safety

- Light scanning sends the selected target-window screenshot to the configured
  endpoint.
- Effect learning sends target-window before/after images and crops of related
  new windows. Full desktop snapshots remain local in the vault.
- API keys are read from environment variables and are not written into the
  knowledge vault.
- PIDs, native window handles, and absolute desktop coordinates are runtime
  observations only and are not persisted as application knowledge.
- Inferred control meanings remain separate from locator verification and
  observed operation effects. Use `teach` to correct a meaning; it cannot bypass
  failed location checks.
- Existing reliable controls are retained when a visual-first rescan omits them.

The same workflow is exposed through `get_ui_learning_readiness`,
`get_ui_learning_status`, `scan_window_knowledge`, `show_ui_controls`,
`list_ui_knowledge_apps`, `search_ui_knowledge`, `search_ui_knowledge_batch`,
`list_ui_commands`,
`run_ui_command`, `run_ui_commands`, `learn_ui_command_effect`,
`explore_ui_workflows`, `list_ui_interactions`, `teach_ui_control`, and
`rebuild_ui_knowledge_index`.

`--full-setup-codex` installs or updates both bundled skills. Use
`--install-codex-skills` when only the skills need to be refreshed.

For MCP client configuration, start the server with `python -m easy_uiauto.mcp.server`.
Control-vector persistence is optional. To enable it, set
`EASY_UIAUTO_CONTROL_VECTOR_DB_DIR` to a directory containing
`control_vector_store.py`; otherwise capture tools still return records but do not persist them.

Control lookup uses the library's canonical `LOCATION` object rather than a
flat selector. Obtain it from a recorded action or from
`get_control_at_position`, then pass the returned `LOCATION` object directly to
`find_control`:

```json
{
  "WindowName": "My Application",
  "Name": "Save",
  "ClassName": "ButtonClass",
  "ControlType": "ButtonControl",
  "foundIndex": 1,
  "AutomationId": "saveButton",
  "Xpath": [
    {"ControlType": "WindowControl", "Name": "My Application", "searchDepth": 1},
    {"ControlType": "ButtonControl", "Name": "Save", "foundIndex": 1, "searchDepth": 2}
  ],
  "Img": "",
  "PARAMETERS": {}
}
```

`find_control` also accepts a complete recorded action containing `LOCATION`
and the complete result from `get_control_at_position`. Legacy flat arguments
remain supported for compatibility, but full XPath data is more reliable for
duplicate or deeply nested controls.

## 🧰 Python API Examples

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

## 🏗️ Project Structure

```text
easyautomation
├── .github/
│   └── workflows/
│       ├── ci.yml
│       ├── publish.yml
│       └── release.yml
├── src/
│   └── easy_uiauto/
│       ├── __init__.py
│       ├── ctrl.py          # Core controller (mouse/keyboard actions)
│       ├── draw.py          # Visual feedback (control highlighting)
│       ├── record.py        # Action recording
│       ├── utils.py         # Control location and utilities
│       ├── mcp/             # MCP server, scanner, vault, and UI CLI
│       └── skills/          # Bundled Codex learning and operation skills
├── tests/
├── CHANGELOG.md
├── LICENSE
├── README.md
├── README.zh-CN.md
└── pyproject.toml
```

## 🚀 Release Automation

The repository uses GitHub Actions for validation and publishing:

- **CI** tests Python 3.11, 3.12, and 3.13, then builds the package.
- **Tagged releases** build and publish through PyPI Trusted Publishing.
- **Build artifacts** include both an sdist and a wheel.

## 🛠️ Development

```bash
pip install -e .[dev]
pytest
ruff check .
```

## 📚 More Examples

For more examples, please refer to the test files in the `demo/` directory or check the docstrings in the source code.

## 📄 License

MIT License. See [LICENSE](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE).

## 💬 Contact

Scan the QR code to add me on WeChat:

![WeChat QR Code](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/or_code.bmp)
