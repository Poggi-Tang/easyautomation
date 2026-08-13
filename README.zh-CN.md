# easy-uiauto
[English](https://github.com/Poggi-Tang/easyautomation/blob/main/README.md) | 简体中文

[![PyPI](https://img.shields.io/pypi/v/easy_uiauto?cacheSeconds=300)](https://pypi.org/project/easy-uiauto/)
[![Python](https://img.shields.io/pypi/pyversions/easy-uiauto?cacheSeconds=300)](https://pypi.org/project/easy-uiauto/)
[![License](https://img.shields.io/github/license/Poggi-Tang/easyautomation)](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE)
[![CI](https://github.com/Poggi-Tang/easyautomation/actions/workflows/ci.yml/badge.svg)](https://github.com/Poggi-Tang/easyautomation/actions/workflows/ci.yml)
[![Publish](https://github.com/Poggi-Tang/easyautomation/actions/workflows/publish.yml/badge.svg)](https://github.com/Poggi-Tang/easyautomation/actions/workflows/publish.yml)

![logo](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/easy-uiauto.png)

**把 Windows 软件界面转换成可检查、可复用的 UI 命令。**

`easy-uiauto` 同时提供 Python 自动化库、MCP 服务、动作录制和 Obsidian
兼容的 UI 知识库。它可以通过 Windows UI Automation 定位控件，学习当前可见
软件中真正有用的区域和操作，并生成类似
`main.keypad.enter-digit-6.click` 的可执行命令。

适用于 Windows 自动化测试、RPA、Agent 工具和桌面软件操作知识沉淀。
项目仅支持 Windows，目前仍处于 `0.x` 阶段。

> ⚠️ 视觉扫描和流程探索仍在持续完善。涉及发送消息、支付、删除、上传、
> 账号变更或其他会产生外部影响的操作前，请先检查生成的命令和风险标记。

## ✨ 包含的能力

- 🖱️ **桌面自动化**：点击、双击、右键、拖拽、滚动、键盘输入、组合键、
  窗口激活和文本填写。
- 🎬 **动作录制**：录制真实操作、高亮当前控件，并生成包含完整 `LOCATION`
  的可复用 Python 动作。
- 🔌 **MCP 服务**：在 Codex、Claude Code 或其他 MCP 客户端中调用同一套能力。
- 🔎 **视觉优先学习**：先从截图中找出关键区域和控件，再映射到本地稳定的
  UIA 控件；默认不遍历完整控件树。
- 🧭 **软件 UI CLI**：把验证通过的控件转换成可搜索的软件命令，执行任务时
  不必重新猜坐标。
- 🖼️ **控件框选反馈**：使用一个点击穿透覆盖层为整页控件编号，保存页面标注图，
  并在执行命令前突出实际目标。
- 🧪 **操作响应学习**：比较操作前后截图，检查变化区域、UIA 属性和新窗口，
  保存有证据支撑的成功条件。
- 🗂️ **可维护知识库**：页面、控件、截图、交互和隔离记录都使用
  Markdown/YAML/PNG，可直接用 Obsidian 或 Git 管理。
- 📐 **稳定记录**：持久化结构化 `LOCATION` 和窗口内归一化视觉提示，不保存
  PID、窗口句柄或桌面绝对坐标。
- 🛟 **分层定位兜底**：依次使用 `LOCATION`、多状态控件图、本地 OCR，最后才
  使用显式开启的远程视觉定位。

## 🚦 能力成熟度

| 能力 | 状态 | 说明 |
| --- | --- | --- |
| Python 鼠标、键盘、窗口和 UIA API | Stable | 原有库能力 |
| 动作录制与回放 | Stable | 生成可复用的结构化动作 |
| MCP 工具与客户端配置 | Beta | 内置 Codex 和 Claude Code 安装命令 |
| 视觉优先扫描与软件 UI CLI | Beta | 需要 OpenAI 兼容视觉接口 |
| 操作前后响应学习 | Beta | 保存本地证据和成功条件 |
| 递归流程探索 | Experimental | 仅探索可恢复操作，并限制递归深度 |

自动流程探索暂不包含滚动和拖拽；基础自动化 API 和直接 MCP 工具仍可显式执行
这两类操作。

## 📦 安装

运行环境：Windows，Python 3.11 或更高版本。

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

## 🔌 MCP 配置

通过 MCP 客户端使用时，安装可选的 MCP 依赖：

```bash
pip install "easy-uiauto[mcp]"
```

需要本地 OCR 与图片模板匹配兜底时，安装视觉可选依赖：

```bash
pip install "easy-uiauto[mcp,vision]"
```

`vision` 包含 OpenCV 模板匹配和 Python Tesseract 适配器。OCR 还需要在系统中安装
Tesseract 可执行文件及相应语言数据，例如 `eng` 或 `chi_sim`。对应 MCP 工具为
`find_control_by_image`、`click_by_image`、`find_text_on_screen`、
`click_text_on_screen`。

远程视觉定位不会运行本地模型，也不依赖厂商 SDK。设置一个 OpenAI 兼容视觉 API 的
端点和凭证后，可使用 `find_control_by_vision` 或 `click_by_vision`：

```bash
EASY_UIAUTO_VISION_API_URL=https://your-api.example/v1/chat/completions
EASY_UIAUTO_VISION_API_KEY=your-api-key
EASY_UIAUTO_VISION_MODEL=your-vision-model
```

这两个工具仅在调用时向配置的端点上传当前截图。应优先采用 UIA、OCR 或图片模板匹配，
多模态 API 只作为最后兜底。

MCP 服务随库一起安装，直接复用同一套自动化 API：

```bash
easy_uiauto --help
easy_uiauto --version
easy_uiauto
```

配置命令通过客户端自身的 CLI 安装、查看或卸载 MCP 条目。基础安装命令遇到同名配置
时不会覆盖。

完整部署 Codex：

```bash
pip install --upgrade "easy-uiauto[mcp,vision]"
easy_uiauto --full-setup-codex \
  --vision-url https://your-api.example/v1/chat/completions \
  --vision-model your-vision-model
```

完整配置会通过隐藏输入询问缺失的 API Key，补齐 Python 视觉依赖和 Tesseract，
仅替换 `easy_uiauto` Codex 条目，安装两个配套 Skill，然后检查 UIA、本地 OCR 和
远程视觉接口。验收使用程序生成的图片，不会上传当前桌面。
当前版本的 MCP 会在每次调用时读取最新视觉配置。安装或更新 Python 包、MCP 条目、
Skill 后只需重启 Codex 一次；如果只是修改视觉配置，不需要再次重启。

需要远程视觉的 Codex 最小部署可使用快速配置命令。API Key 未存在于 Windows
用户环境变量时，命令通过终端隐藏输入；由非交互式 Agent 执行时则弹出密码框。
随后保存三个视觉变量，仅替换 `easy_uiauto` MCP 条目，并跳过 OCR 安装和 UI 实测：

```bash
easy_uiauto --quick-setup-codex \
  --vision-url https://your-api.example/v1/chat/completions \
  --vision-model your-vision-model
```

也可以直接对 Codex 说：

> 从 PyPI 安装最新版 `easy-uiauto[mcp,vision]`，然后执行
> `easy_uiauto --full-setup-codex --vision-url URL --vision-model MODEL`；不要检查无关项目，
> 不要搜索网页，报告每项验收结果和耗时，最后提示我只重启 Codex 一次。不要仅因视觉
> 环境变量发生变化要求再次重启。

```bash
easy_uiauto --install-codex
easy_uiauto --show-codex-config
easy_uiauto --uninstall-codex

easy_uiauto --install-claude-code
easy_uiauto --show-claude-code-config
easy_uiauto --uninstall-claude-code
```

Codex 会写入全局 `config.toml`；Claude Code 会写入 `user` 作用域，因此对本机所有
项目生效。安装、更新或卸载后需要重启相应客户端。

也提供常驻 TCP 服务：

```bash
easy_uiauto_service --help
python -m easy_uiauto.mcp.service --port 9876
```

## 🧭 学习一个软件

安装包中包含两个 Codex Skill：

- `$easy-uiauto-learning`：扫描和探索当前可见的 Windows 软件。
- `$easy-uiauto-operate`：使用已有知识库中的命令操作软件。

完成配置并重启 Codex 一次后，可以直接这样说：

```text
使用 $easy-uiauto-learning，学习当前打开的微信界面，生成微信的控件知识库和 UI 命令。
```

### 🔎 轻度扫描

轻度扫描只观察当前可见页面，不会点击软件中的控件。

```bash
easy_uiauto_ui scan "窗口标题"
easy_uiauto_ui apps
easy_uiauto_ui commands <app-id>
easy_uiauto_ui show <app-id>
```

默认的 `visual-first` 策略会把目标窗口截图发送到已配置的视觉接口，找出关键区域和控件，
再把坐标映射到本地 UIA 控件。最终保存稳定 `LOCATION`、控件图、功能含义、同义词、
风险和生成的命令。支持位于副屏上的窗口。扫描还会保存 `<页面-id>.annotated.png`，
并在当前窗口短暂绘制相同编号的控件框；不需要实时框选时可传入 `--no-overlay`。

通过 MCP 学习时，每 5 秒会报告任务 ID、当前阶段、已耗时和控件验证进度。如果客户端
停止等待，应先调用 `get_ui_learning_status`；状态为 `running` 表示原扫描仍在 MCP 进程中
继续执行，不能重复启动扫描。只有 `failed` 才表示学习失败。
普通“学习/扫描”请求到基础扫描结束即完成；只有明确要求“深度学习”或“自主交互探索”时，
才会执行 `explore_ui_workflows`。

视觉定位不适用或需要诊断完整 UIA 覆盖时，可使用 `--strategy full-uia`。该模式会遍历
可见 UIA 树并额外执行语义分析，因此更慢，产生的记录也更多。

### 🧪 深度探索

深度探索会先执行轻度扫描，再操作已知且可恢复的命令，并记录操作造成的变化。
遇到新页面或弹窗时，可以在限制深度内继续学习。

```text
使用 $easy-uiauto-learning，深度学习当前打开的微信，采用 safe 策略，最多操作 30 次，
递归深度 3。不要发送消息、删除数据、退出登录，也不要执行任何需要确认的操作。
```

```bash
easy_uiauto_ui explore <app-id> --policy safe --max-actions 30 --max-depth 3
easy_uiauto_ui interactions <app-id>
```

每次操作都会保存前后截图、局部像素差异、发生变化的 UIA 属性、新增或短暂出现的窗口、
观察到的成功条件和恢复结果。探索器会按“页面+命令”去重；恢复失败或检测到用户和其他
窗口干扰时立即停止。

`safe` 只执行标记为安全的命令；`supervised` 还允许可恢复的状态变更命令。
外部影响、破坏性和需要确认的命令不会被自动探索。探索器暂不实现滚动和拖拽。

### ▶️ 执行已学习命令

```bash
easy_uiauto_ui search <app-id> "检索词"
easy_uiauto_ui show <app-id> --include executable
easy_uiauto_ui run <app-id> <页面.区域.控件.动作> --text "可选文本"
easy_uiauto_ui batch <app-id> '["main.keypad.enter-digit-6.click", "main.keypad.enter-addition-operation.click"]'
easy_uiauto_ui learn-effect <app-id> <command> --recover
easy_uiauto_ui teach <app-id> <control-id> "控件含义" intent "功能说明"
```

运行时依次使用 `LOCATION` → 多状态控件图 → 本地 OCR → 显式开启的远程视觉。
最后一级需要传入 `--allow-vision-fallback` 或 `allow_vision_fallback=true`。
控件失效、缺失或匹配不唯一时会进入隔离区，不会继续点击。单条和批量命令默认会用
红框预览实际解析出的目标；可通过 `--no-highlight` 或 `highlight=false` 关闭约 100 ms
的可视预览等待。

批量执行只适用于同一个稳定页面，发生页面跳转后应拆成下一批。涉及外部影响或破坏性的
命令必须传入 `--confirm` 或 `confirm=true`。

### 🗂️ 知识库结构

默认知识库目录为 `~/easy_uiauto_vault`，也可通过
`EASY_UIAUTO_KNOWLEDGE_DIR` 指定其他位置。Markdown/YAML 和 PNG 是数据源：

```text
applications/<app-id>/
├── pages/                 页面记录
├── regions/               功能区域
├── controls/              已验证和已观察控件
├── interactions/          操作前后响应记录
├── images/                页面、控件和交互截图
├── quarantine/            失效或匹配不明确的控件
└── operations/UI-CLI.md   自动生成的命令目录
```

`.easy_uiauto/index.json` 只是可删除的检索缓存。手工修改 Markdown 后，使用
`easy_uiauto_ui reindex <app-id>` 重建即可。

控件和区域只保存窗口内的 `normalized_rect`，并明确标记
`geometry_role: visual-hint-only`。它只用于页面标注和视觉对比，不能作为控件身份，
也不能直接用于点击。实时框选和操作会先重新解析当前 `LOCATION`，失败后才使用唯一
图片模板或本地 OCR 兜底。0.6 版重建旧知识库索引时，会自动清除历史绝对坐标、PID
和窗口句柄。

### 🔐 隐私与安全

- 轻度扫描会把选中的目标窗口截图发送到已配置的视觉接口。
- 操作响应学习会发送目标窗口的前后图，以及与操作相关的新窗口裁剪图；完整桌面快照
  只保存在本地知识库。
- API Key 从环境变量读取，不会写入知识库。
- PID、原生窗口句柄和桌面绝对坐标只作为本次运行观测，不会写成长期软件知识。
- 推断出的控件含义、定位验证和实际操作响应分别记录，不会混为“已经验证”。可以使用
  `teach` 修正含义，但人工修正不能绕过失败的定位检查。
- 视觉优先重扫偶尔漏掉控件时，已有可靠控件不会被直接删除。

对应 MCP 工具包括 `get_ui_learning_readiness`、`get_ui_learning_status`、
`scan_window_knowledge`、`show_ui_controls`、
`list_ui_knowledge_apps`、
`search_ui_knowledge`、`list_ui_commands`、`run_ui_command`、`run_ui_commands`、
`learn_ui_command_effect`、`explore_ui_workflows`、`list_ui_interactions`、
`teach_ui_control` 和 `rebuild_ui_knowledge_index`。

`--full-setup-codex` 会安装或更新两个配套 Skill。只需要更新 Skill 时，使用
`--install-codex-skills`。

MCP 客户端配置可使用 `python -m easy_uiauto.mcp.server` 启动服务。
控制向量持久化是可选能力：将 `EASY_UIAUTO_CONTROL_VECTOR_DB_DIR` 指向包含
`control_vector_store.py` 的目录即可启用；未配置时，采集工具仍会返回记录，但不会持久化。

控件查找使用库原生的完整 `LOCATION` 结构，而不是扁平选择器。该结构可从录制动作中
取得，也可调用 `get_control_at_position` 通过坐标生成；随后将返回的 `LOCATION` 对象
直接传给 `find_control`：

```json
{
  "WindowName": "我的应用",
  "Name": "保存",
  "ClassName": "ButtonClass",
  "ControlType": "ButtonControl",
  "foundIndex": 1,
  "AutomationId": "saveButton",
  "Xpath": [
    {"ControlType": "WindowControl", "Name": "我的应用", "searchDepth": 1},
    {"ControlType": "ButtonControl", "Name": "保存", "foundIndex": 1, "searchDepth": 2}
  ],
  "Img": "",
  "PARAMETERS": {}
}
```

`find_control` 也可直接接收包含 `LOCATION` 的完整录制动作，或者
`get_control_at_position` 的完整返回结果。旧的扁平参数仍兼容，但遇到重名或深层控件时，
应使用包含完整 XPath 的结构。

## 🧰 Python API 示例

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

## 🏗️ 项目结构

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
│       ├── ctrl.py          # 核心控制器（鼠标/键盘操作）
│       ├── draw.py          # 视觉反馈（控件高亮）
│       ├── record.py        # 动作录制
│       ├── utils.py         # 控件定位和通用工具
│       ├── mcp/             # MCP 服务、扫描器、知识库和 UI CLI
│       └── skills/          # 随包提供的 Codex 学习与操作 Skill
├── tests/
├── CHANGELOG.md
├── LICENSE
├── README.md
├── README.zh-CN.md
└── pyproject.toml
```

## 🚀 发布自动化

仓库通过 GitHub Actions 完成检查和发布：

- **CI**：使用 Python 3.11、3.12 和 3.13 运行测试，并检查构建结果。
- **Tag 发布**：通过 PyPI Trusted Publishing 构建并上传版本，无需维护 PyPI Token。
- **构建产物**：同时生成 sdist 和 wheel。

## 🛠️ 本地开发

```bash
pip install -e .[dev]
pytest
ruff check .
```

## 📚 更多示例

更多示例请参考 `demo/` 目录中的测试文件或查看源代码中的文档字符串。

## 📄 许可证

MIT License，详见 [LICENSE](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE)。

## 💬 联系方式

扫描以下二维码添加我的微信：

![微信二维码](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/or_code.bmp)
