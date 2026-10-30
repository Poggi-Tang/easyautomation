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

图片标注默认并推荐使用 **YOLO＋RapidOCR**：输入图片路径和输出目录，即可生成文字标注图、
图形控件标注图、合并标注图、坐标、识别证据和本地 HTML 查看页。

> 视觉扫描和流程探索仍在持续完善。涉及发送消息、支付、删除、上传、
> 账号变更或其他会产生外部影响的操作前，请先检查生成的命令和风险标记。

## 包含的能力

- **桌面自动化**：点击、双击、右键、拖拽、滚动、键盘输入、组合键、
  窗口激活和文本填写。
- **动作录制**：录制真实操作、高亮当前控件，并生成包含完整 `LOCATION`
  的可复用 Python 动作。
- **MCP 服务**：在 Codex、Claude Code 或其他 MCP 客户端中调用同一套能力。
- **视觉优先学习**：先从截图中找出关键区域和控件，再映射到本地稳定的
  UIA 控件；默认不遍历完整控件树。
- **状态语义理解**：区分稳定功能和动态值，保留联系人名称、头像等实体，
  建立控件关系，并解释“消息为空，因此发送按钮禁用”这类可见状态。
- **软件 UI CLI**：把验证通过的控件转换成可搜索的软件命令，执行任务时
  不必重新猜坐标。
- **控件框选反馈**：使用一个点击穿透覆盖层为整页控件编号，保存页面标注图，
  并在执行命令前突出实际目标。
- **本地图片标注**：默认 CPU 运行 YOLO＋RapidOCR，无需视觉 API；蓝色 G 编号表示图形、
  橙色 T 编号表示文字，使用 1 像素边框和贴边标签，支持几何纠正及小文字自动补检。
- **操作响应学习**：比较操作前后截图，检查变化区域、UIA 属性和新窗口，
  保存有证据支撑的成功条件。
- **可维护知识库**：页面、控件、截图、交互和隔离记录都使用
  Markdown/YAML/PNG，可直接用 Obsidian 或 Git 管理。
- **稳定记录**：持久化结构化 `LOCATION` 和窗口内归一化视觉提示，不保存
  PID、窗口句柄或桌面绝对坐标。
- **分层定位兜底**：依次使用唯一 `AutomationId`、上下文 XPath、稳定属性组合、
  多状态控件图、本地 OCR，最后才使用显式开启的远程视觉定位。
- **快速定位锚点**：优先使用窗口内唯一的 `AutomationId`，缺失或重复时再
  使用 XPath 层级消歧。

## 能力成熟度

| 能力 | 状态 | 说明 |
| --- | --- | --- |
| Python 鼠标、键盘、窗口和 UIA API | Stable | 原有库能力 |
| 动作录制与回放 | Stable | 生成可复用的结构化动作 |
| MCP 工具与客户端配置 | Beta | 内置 Codex 和 Claude Code 安装命令 |
| 视觉优先扫描与软件 UI CLI | Beta | 需要 OpenAI 兼容视觉接口 |
| 本地图片标注 | Experimental | 默认并推荐 YOLO＋RapidOCR，无需 API key；MCP 和 CMD 支持普通界面、控制框图和电气元件图 |
| 操作前后响应学习 | Beta | 保存本地证据和成功条件 |
| 递归流程探索 | Experimental | 仅探索可恢复操作，并限制递归深度 |

自动流程探索暂不包含滚动和拖拽；基础自动化 API 和直接 MCP 工具仍可显式执行
这两类操作。

## 安装

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

只需本地图片标注时，安装 `image` 扩展并运行：

```cmd
python -m pip install "easy-uiauto[image]"
easy_uiauto_image "C:\images\input.bmp" "C:\images\results"
```

每次在输出目录下新建独立结果子目录，生成 `text.png`、`controls.png`、`combined.png`、
`elements.json`、`elements.csv`、诊断记录和 `index.html`。结束后打开 `index.html`
集中查看。模型在本机 CPU 运行，首次使用会下载并校验 YOLO 权重。默认 `ui` 用于普通软件
截图，控制框图和电气模型图统一加 `--profile block-diagram`。详见
[CMD 本地图片标注](#cmd-本地图片标注)。

## MCP 配置

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

普通 Python 代码只需安装本地视觉依赖并直接调用矩形分割器，不会启动 MCP、调用 OCR
或上传数据：

```bash
pip install "easy-uiauto[vision]"
```

```python
from easy_uiauto.visual import segment_screen_rect

result = segment_screen_rect(
    {"left": 100, "top": 200, "right": 900, "bottom": 800},
    min_width=6,
    min_height=6,
    save_annotated=True,
)
for box in result["boxes"]:
    print(box["rect"], box["parent_id"])
print(result["relations"], result["groups"])
```

如果 UIA 只暴露一个大画布、内部元素没有 UIA 节点，可把控件返回结果或其中的矩形
直接交给本地矩形检测器：

```text
get_control_at_position(x, y)
detect_visual_elements(rect=<上一步返回的控件对象>)
```

`detect_visual_elements` 只截图传入的矩形，返回编号后的无类别候选框、屏幕绝对矩形、
相对输入区域的归一化矩形、包含关系、置信度、耗时和标注预览图路径。本地布局图还会
生成 `right-of`、`below`、边缘对齐、投影重叠和自适应间距关系，并据此生成尺寸可变的
组合框。组合签名不包含元素宽度、高度和当前内容，因此名称变长、消息换行或内容增高后
仍属于相同结构；垂直重复列表项和相邻表格行不会被误合并。它不会遍历 UIA、运行 OCR、
理解控件含义、下载模型或调用远程 API。

远程视觉定位不会运行本地模型，也不依赖厂商 SDK。设置一个 OpenAI 兼容视觉 API 的
端点和凭证后，可使用 `find_control_by_vision` 或 `click_by_vision`：

```bash
EASY_UIAUTO_VISION_API_URL=https://your-api.example/v1/chat/completions
EASY_UIAUTO_VISION_API_KEY=your-api-key
EASY_UIAUTO_VISION_MODEL=your-vision-model
```

这两个工具仅在调用时向配置的端点上传当前截图。应优先采用 UIA、OCR 或图片模板匹配，
多模态 API 只作为最后兜底。

### 图形配置窗口

Windows x64 安装包的 wheel 和源码包均内置 `easy_uiauto/bin/easy-uiauto-config.exe`。
安装后运行 `easy_uiauto_config` 即可打开内置 EXE；也可直接双击独立构建产物
`dist/easy-uiauto-config.exe`。窗口用于配置知识库路径、API URL、API Key 和 API Model；
EXE 自带运行环境，单独运行不需要另装 Python。API Key 按明文显示，
Model 下拉框既可输入，也可点击右侧“获取”后选择服务提供的模型。“浏览”用于选择
本地知识库目录，“保存配置”将四项设置写入 Windows 当前用户环境。

API URL 支持服务根地址、`/v1` 地址或完整的 `/chat/completions` 地址；保存时转换为
现有 MCP 使用的完整请求地址。获取模型列表不会修改已保存配置，也不会发送截图。
新版本 MCP 在下次调用时读取知识库路径和 API 配置；旧版 MCP 需先升级并重启。
修改知识库路径不会自动迁移原有内容。本窗口仅配置一个 API。

从源码构建 EXE（Windows）：

```cmd
python -m pip install pyinstaller
python scripts/build_config_exe.py --bundle-package
python -m build
python scripts/verify_config_bundle.py
```

打包前须先生成 EXE；`--bundle-package` 会将 EXE 复制进包资源目录。
仓库 CI 和标签发布流程会先在 Windows 编译 EXE，再生成并检查 wheel、源码包，
确认两者都包含同一份 EXE 后才允许发布。编译产物通过构建工件传递，不提交到 Git。

### 窗口边缘裁剪

学习、前后截图和回放校验会先排除 Windows 的不可见缩放边框，再用本地像素矩阵检测
接近整窗的闭合矩形。外围安静时检查是否有被遗漏的文字或控件；外围是桌面或其他窗口时，
额外要求连续外框、近似对称的边距和横贯窗口的标题栏色带变化。边界不确定则保留原图。
保留可见标题栏和边框，裁剪后的截图原点也会同步调整。已有知识建议重扫以更新图像基准。

MCP 使用 `take_window_screenshot(window_name="新建项目", auto_crop=True)`，同时返回图片、
`capture_bounds`、`crop_offset` 和尺寸；`auto_crop=False` 保留原始 UIA 外框。
原有 `take_screenshot` 的全屏和指定区域行为保持不变。

已有图片可直接使用本地算法（需安装 `easy-uiauto[vision]`）：

```python
from PIL import Image
from easy_uiauto.visual import crop_window_image

result = crop_window_image(Image.open("window.png"), max_margin=64)
result["image"].save("window-cropped.png")
print(result["cropped"], result["rect_in_input"], result["offset"])
```

### 命令行配置

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


### 动作结果与等待验证

桌面动作现在返回包含 `ok`、`code`、`verified` 的统一 JSON；旧文本位于 `message`。
`ok=true` 表示动作执行成功，后续用 `wait_for_control`、`wait_for_text`、
`get_window_state` / `wait_for_window` 验证预期效果。新增 `get_capability_readiness`
分别检查 UIA、OCR、图片匹配和远程视觉配置。滚动优先使用 `scroll`，
`mouse_scroll` 是相同参数的兼容别名。

参数示例、错误代码和迁移说明见 [MCP 动作结果与等待](docs/mcp-results-and-waits.md)。

## 学习一个软件

安装包中包含两个 Codex Skill：

- `$easy-uiauto-learning`：扫描和探索当前可见的 Windows 软件。
- `$easy-uiauto-operate`：使用已有知识库中的命令操作软件。

完成配置并重启 Codex 一次后，可以直接这样说：

```text
使用 $easy-uiauto-learning，学习当前打开的微信界面，生成微信的控件知识库和 UI 命令。
```

### 轻度扫描

轻度扫描只观察当前可见页面，不会点击软件中的控件。

```bash
easy_uiauto_ui scan "窗口标题"
easy_uiauto_ui apps
easy_uiauto_ui commands <app-id>
easy_uiauto_ui show <app-id>
```

默认的 `visual-first` 策略会把目标窗口截图发送到已配置的视觉接口，大图按原始分辨率
分成重叠区域分析，逐项识别可见控件和有意义的状态，
再把坐标映射到本地 UIA 控件。最终保存稳定 `LOCATION`、控件图、功能含义、同义词、
风险和生成的命令。支持位于副屏上的窗口。扫描还会保存 `<页面-id>.annotated.png`，
并在当前窗口短暂绘制相同编号的控件框；不需要实时框选时可传入 `--no-overlay`。

桌面扫描会额外枚举 Explorer 暴露的桌面图标，补齐视觉漏识别的快捷方式和文件；返回的
`coverage` 包含分块数、桌面项数和补充数。超过扫描上限时会明确返回 `truncated=true`。
桌面项使用原生清单并通过屏幕命中检查排除被窗口遮挡的图标，视觉接口只分析任务栏，
避免重复生成大量图标描述。`desktop_items_occluded` 单独报告遮挡项数量。
共享 UIA 容器的不同视觉元素会分别保存和标注；只有取得独立且验证通过的定位后才会生成
可执行命令。没有独立定位的元素保留为观察记录。仅凭 UIA 名称补充的桌面项不宣称已验证
打开后的效果，并保留未知风险及确认标记。

`get_control_tree` 通过 Win32 选择可见窗口，只为该句柄绑定 UIA，使用独立工作进程，
不会激活窗口或改变最大化状态。默认 `max_depth=5`（最多 10）、`max_nodes=500`
（最多 5000）、`timeout_seconds=15`（0.1～60 秒，含进程启动）。达到深度或节点上限后
停止枚举子节点。正常树的 `_inspection` 记录达到的上限和提供程序错误；底层 UIA 阻塞时
返回 `ok=false`、`code="timeout"`，`data` 中包含最近的部分树/检查点，并停止本次工作
进程。不会终止目标软件。`get_ui_learning_status` 只跟踪学习任务，不跟踪这个树查询。

通过 MCP 学习时，每 5 秒会报告任务 ID、当前阶段、已耗时和控件验证进度。如果客户端
停止等待，应先调用 `get_ui_learning_status`；状态为 `running` 表示原扫描仍在 MCP 进程中
继续执行，不能重复启动扫描。只有 `failed` 才表示学习失败。
普通“学习/扫描”请求到基础扫描结束即完成；只有明确要求“深度学习”或“自主交互探索”时，
才会执行 `explore_ui_workflows`。

运行时定位验证失败的控件会持续隔离，重扫、重建索引和语义教学都不会解除隔离。修正
定位或恢复所需可见页面后，调用 `revalidate_ui_control(app_id=..., control_id=...)`
做显式复核：只检查 LOCATION/图像，不执行控件动作；OCR 单独成功不足以解隔离，原失败
证据仍保留。自动探索会跳过已经尝试过的控件/动作，包括无可见变化和受干扰的记录；
需要主动重复观察时仍可显式调用 `learn_ui_command_effect`。激活窗口保持最大化状态；
前台被其他应用抢走后，停止后续页面发现和恢复输入。

有明确原生文字的菜单、按钮、复选框和单选按钮会核对 UIA 名称与视觉语义，处理
中文/英文常见动作、助记键
和快捷键后缀；例如“全选”不能绑定为“转到”，“关闭”不能绑定为“最大化”。冲突的视觉
观察保留为诊断，不生成相应错误命令。视觉优先扫描额外读取目标窗口内有界的浅层
菜单/标题栏 UIA 元素，用原生文字补充独立动作记录，不枚举桌面 UIA 根节点。
执行前重复检查已知语义和当前动作文字；明确冲突不能被图像/OCR 兜底覆盖，旧错误记录
也不会因重建索引而重新变成可执行命令。该检查覆盖常见动作名称，不把未知自绘控件的
模型推断当作功能验证。自动探索只选择本轮实际观察到的控件，并跳过禁用项以及不能用
Escape 撤销的新建文档/窗口；子探索停止后，父层也不继续发送恢复按键。
“查找内容”“替换为”等已知标签的输入框可使用独立原生标签证据；编辑器中的动态文字
不参与动作身份推断。主窗口被禁用时，恢复仅激活经所有权确认的模态窗口，不对主窗口
盲发按键；学习所属弹窗时保留应用主窗口名称。递归返回时会重新检查动作是否已尝试。

客户端调用超时不代表后台任务停止，应先查 `get_ui_learning_status`；杀死 MCP 进程会
终止工作线程并丢失内存任务状态。知识库文件使用 UTF-8，在 PowerShell 查看时可使用
`Get-Content -Encoding UTF8`，避免把终端解码乱码误认为文件损坏。

视觉定位不适用或需要诊断完整 UIA 覆盖时，可使用 `--strategy full-uia`。该模式会遍历
可见 UIA 树并额外执行语义分析，因此更慢，产生的记录也更多。

### 图片文件识别与标注

调用 MCP 工具 `analyze_ui_image(image_path="C:/images/screenshot.png")` 可直接分析
PNG、JPEG、WebP 等图片，无需打开软件或操作桌面。**默认 `engine="local"` 使用本地
YOLO＋RapidOCR，也是当前推荐组合。** 在本项目 `test1/test2/test3` 软件截图与视觉 API、
Tesseract＋tessdata_best 的实测对比中，这个组合的文字覆盖和整体标注表现更好；这是本次
样例的结论，其他界面仍需检查漏检、误框和错字。

图形检测采用微软官方 [OmniParser v2.0 的 icon_detect_v3 权重](https://huggingface.co/microsoft/OmniParser-v2.0/tree/main/icon_detect_v3)
`model.pt`，架构为 **YOLOv9-E**；文字识别采用 RapidOCR 的 **PP-OCRv6 small ONNX**。
这套权重没有针对本项目的电路图或 Simulink 图片重新训练。安装本地依赖：

```cmd
python -m pip install ".[mcp,image]"
```

本地模式无需 API URL、Key 或远程模型；首次使用会下载并校验 YOLO 权重，之后复用缓存。
本地依赖缺失时返回安装错误，不会自动上传图片。需要比较远程视觉模型时，显式调用
`analyze_ui_image(image_path="C:/images/screenshot.png", engine="api")`，该模式才读取
已配置的视觉 API。`model`、`analysis_scale`、`analysis_context` 仅适用于 API 模式。
MCP 本地模式默认 `profile="ui"`，控制框图和电气模型图统一选 `block-diagram`。
独立的 `circuit` 配置已移除，旧调用请改用 `block-diagram`，已有标注的模型图修正也使用该名称。
YOLO 默认置信度阈值为 **0.1**；MCP 可传 `confidence=0.1`，CMD 可用 `--confidence` 调整。

本地标注默认再对图形框做**矩形轮廓吸附**：灰度阈值 210，候选宽 30～600、高 12～60
像素，选择 IoU 最高且至少为 0.35 的轮廓，并要求候选宽、高均不超过原框的 1.6 倍。
额外检查轮廓面积，减少把文字笔画当成矩形；不匹配时保持原框。它用于补全输入框、按钮等
闭合矩形边框，不替代任意形状控件的检测。文字框和编号保持不变，`detected_rect` 保留
原始预测，`corrected_rect`、`snap_iou` 和 `geometry_correction` 记录吸附结果；`rect`、
`bbox`、中心点和归一化坐标同步更新。本地默认保留 1 像素边距。
MCP 的 `rectangle_snap=False` 或 CMD 的 `--no-rectangle-snap` 可关闭此步骤。

两种模式默认先用本地像素算法裁剪可靠的窗口外框，再执行识别；无法确定外框时保留原图。
`auto_crop=False` 可保留完整输入。返回 `crop.offset` 和 `input_image_path`，裁剪后的
坐标加上该偏移即可映射回 EXIF 方向校正后的原图。`image_path` 指向实际识别输入，
裁剪时为保存的 `analysis-input.png`；后续修正使用这个文件，避免坐标错位。
工具返回
`annotated_image` 标注图、`coordinates_json` 坐标文件和编号对应的 `elements`。
每项包含名称、推测功能、置信度、`bbox`（x、y、宽、高）、`rect`（左上边界包含、右下
边界不包含）、中心点和归一化坐标。坐标以实际识别输入 `image_path` 的左上角为原点，
图形使用蓝色框与 `G1…` 标签，文字使用橙色框与 `T1…` 标签，均为未经 UIA 或功能执行
验证的观察项，不写入应用知识库。`annotation_label` 对应显示编号，原始数字 `id` 供
修正工具调用。`output_dir` 指定独立结果文件夹的父目录，`max_elements` 限制返回数量，
超出时明确报告 `truncated`。图片附件需先在 MCP 所在机器上保存为本地文件。
API 模式的小图默认自动放大识别（最高 4 倍），区域和控件坐标均回映识别输入；传入 `analysis_scale=1`
可关闭放大，1 到 4 之间的其他值可指定识别倍率。
`box_padding` 为识别输入的像素边距（本地默认 1，API 默认 2，传 0 关闭）；`detected_rect` 保留原始预测，
`rect` 和 `bbox` 包含边距，并使已知组合控件的整体框包含附属子元素。
放大预览 `annotated_preview_image` 会单独绘制 1 像素边框和紧凑编号，避免把线宽和标签
一起放大。标签始终放在对应框的四边外侧，间隔 2 像素，避让时只沿边调整，不移到远处。
标注图通常保持原尺寸；若图像边缘容不下标签，会增加白色画布边距，原图坐标保持不变。
`annotation_layout.image_origin` 记录原图在画布中的偏移，`labels` 记录各编号标签位置，
`annotation_preview_layout` 记录预览布局。可用 `analysis_context` 指定分析范围和粒度，例如按完整电路模型元件标注，
不把元件内部的电阻、电容等图案拆成独立控件。

框选不完整时，可调用
`correct_ui_image_annotations(coordinates_json=".../elements.json", profile="rectangles")`
对已有结果的全部图形框执行同一套矩形吸附，或用 `refine_ids` 选择指定项；不重新调用
模型、不改文字、不丢失原 JSON 的来源和坐标元数据。它也可修正显式 API 分析的已有结果。

独立图标也可调用
`correct_ui_image_annotations(coordinates_json=".../elements.json", refine_ids=[1,3,9])`
按编号在原图上进行本地像素修正，无需再次调用视觉 API。自动修正适合浅色简单背景上的
独立图标；连接线或背景复杂、无法可靠分离时会跳过并给出原因。也可传入
`corrections=[{"id":3,"rect":{"left":110,"top":65,"right":125,"bottom":81}}]`
直接指定精确边界，右下坐标不包含且不额外加边距。自动修正默认保留 1 像素边距。
默认流程不会把 `icons` 前景拟合应用到所有控件；它可能把输入框缩成内部文字范围，
应只对合适的独立图标显式调用并检查修正前后。
编号、语义、原始 `detected_rect` 和标注样式保持一致，新生成 PNG/JSON，不覆盖旧结果；
`corrected_rect` 与 `correction.updates` 记录修正内容。新识别结果记录原图哈希，防止修正
时误用已变化的图片；旧结果仅校验尺寸。修正结果仍是未经 UIA 验证的观察项。

模型图修正可加 `profile="block-diagram"`：从原图像素提取矩形外框、双圆变压器、带接地符号的
负载支路和粗线母线，匹配所选观察项的几何边界，并归并匹配到同一图形的重复结果。
`discover_missing=true` 可额外补充未匹配的母线，不自动补充其他类型。保留主编号和模型
原始预测；候选数量、归并及新增编号记录在 `correction` 中。该模式不读取旧标注作为
参考坐标，不调用视觉 API；支持上述图形结构，不能保证适用所有电路符号，需检查跳过项。
控制框图还支持完整模块、常量、Mux 外框，From/Goto 尖头标签，电机/仪表椭圆、端口圆环
和接地符号。母线检测排除闭合外框及其阴影边缘。无法确定的预测保持原样，并列入
`correction.needs_review_ids`，不能当作已确认元件。`coverage` 记录分块边界和原始/归并
观察数量；跨分块去重不再要求模型输出的角色、类型完全一致，冲突证据保存在
`tile_observations` 中。
文字也是控件：标签、标题、数值和端口文字用独立的 `text-label` 编号，文字框与图形框
可以重叠。用 `text-of`/`label-of` 表达所属关系，外部文字不扩大图形框。
电路修正保留 `text_control_ids`；文字坐标需要文字检测/OCR 或复核，不能用闭合元件
外框验证，也不能因缺少外框就认定文字是误检。

### CMD 本地图片标注

CMD 同样默认并推荐使用 YOLO＋RapidOCR，不读取已配置的视觉 API。
从仓库运行，不需要启动 Codex 或 MCP 服务：

```cmd
cd /d C:\path\to\easyautomation
annotate-image.cmd "C:\images\input.bmp" "C:\images\results"
```

#### IEEE9 标注示例

使用本地 YOLO＋RapidOCR 标注 [IEEE9.bmp](https://raw.githubusercontent.com/Poggi-Tang/easyautomation/v0.9.10/docs/examples/IEEE9/IEEE9.bmp)：

```cmd
annotate-image.cmd "docs\examples\IEEE9\IEEE9.bmp" "%TEMP%\easy-uiauto-ieee9" --profile block-diagram
```

该示例识别出 **42 个图形、102 个文字区域**。

原图：

![IEEE9 原图](https://raw.githubusercontent.com/Poggi-Tang/easyautomation/v0.9.10/docs/examples/IEEE9/IEEE9.png)

文字与图形合并标注：

![IEEE9 文字与图形合并标注](https://raw.githubusercontent.com/Poggi-Tang/easyautomation/v0.9.10/docs/examples/IEEE9/results/combined.preview.png)

标注结果：[图形](https://raw.githubusercontent.com/Poggi-Tang/easyautomation/v0.9.10/docs/examples/IEEE9/results/controls.png)、
[文字](https://raw.githubusercontent.com/Poggi-Tang/easyautomation/v0.9.10/docs/examples/IEEE9/results/text.preview.png)、
[坐标 JSON](https://github.com/Poggi-Tang/easyautomation/blob/v0.9.10/docs/examples/IEEE9/results/elements.json)。
下载[完整示例](docs/examples/IEEE9)后，打开 `docs/examples/IEEE9/results/index.html` 即可交互查看结果。

前两个参数分别是图片路径和输出父目录，支持中文和空格；路径请加双引号。
每次在输出目录下建立独立的 `图片名-时间-随机编号` 子目录，避免覆盖旧结果。
命令结束后打印该目录和 `index.html` 路径，双击 `index.html` 即可集中查看。

| 文件 | 内容 |
| --- | --- |
| `text.png` | 橙色 T 编号的文字标注 |
| `controls.png` | 蓝色 G 编号的图形控件标注 |
| `combined.png` | 文字与图形合并标注 |
| `original.png` | 应用 EXIF 方向、透明区域铺白后的原图 |
| `analysis-input.png` | 自动裁剪后的实际识别输入；关闭裁剪时为完整输入 |
| `elements.json` / `elements.csv` | 文字内容、编号、坐标、来源、所属关系和待复核信息 |
| `raw-yolo.json` / `raw-ocr.json` | 原始检测与 OCR 输出 |
| `ocr-recovery.json` | 小文字补检候选、过滤原因和两次识别证据 |
| `geometry-candidates.json` | 本地几何候选与证据 |
| `rectangle-snap.json` | 矩形吸附候选、IoU、修正前后坐标和跳过原因 |
| `ui-fusion.json` | 空白过滤、紧凑边框保护、重复图标补检及 OCR 冲突证据 |
| `before-snap.png` / `snap-compare.png` | 吸附前标注和红色原框、绿色修正框、蓝色未变框的对比 |
| `summary.json` / `run.log` | 数量、耗时、成功/失败状态与处理日志 |
| `index.html` | 三张标注图和资源文件的本地查看页 |

默认 `--profile ui` 用于普通软件截图，并与 MCP 本地默认配置一致；控制框图使用
`--profile block-diagram`，电气模型图也使用此配置。UI 模式保留 YOLO 图形候选，
先检查候选框及周围 2 像素是否存在可见前景或边框，过滤纯空白假框；带真实边框的空输入框
仍保留。小方框控件以可见闭合边框为依据，不会仅因 OCR 重叠就删掉图形；框内误读成
`M`、勾选符号等的短 OCR 项记录为冲突，原始 OCR 输出不变。相同小方框的嵌套预测归并，
原始 YOLO 编号保存在 `yolo_proposal_ids`。

UI 模式还从当前图片已检测的小图标提取模板，按相关度至少 0.97 补检相同像素图案；
最多使用 12 个模板，避开已有小控件和正常 OCR 文字，不读取旧标注或参考图坐标。
补检来源为 `local-template-recovery`，`confidence_kind=template-correlation`，该分数
不是 YOLO 置信度，也不验证控件语义。过滤、保护和补检证据均写入 `ui-fusion.json`。
对齐排列的工具栏、停靠按钮候选会保护图标，避免被低置信度 OCR 字母或符号误删，
并保留浅色禁用图标的全部像素部件。重复树形箭头、前缀图标、文字前的搜索放大镜和
成组实心标记可与文字分离；原文字、原坐标和纠正过程均保留，几何拟合不验证控件功能。
表格网格、残缺行和重复列表区域单独记录为布局证据，不再与其内部文字、图标叠画为独立图形。
箭头和关闭符号在像素证据匹配时合并重复 OCR；有邻近区域支持的滚动条滑块，以及被中文标签
末尾误读为 Q/O 的放大镜图形，可在本地分离，原文字和坐标保存在纠正证据中。带浅灰闭合
边框的长输入框和按钮受保护。密集小控件页面另存 2 倍预览，重新绘制细框和相邻编号，
HTML 查看器优先显示该预览；坐标仍相对识别输入图。
UI 图片长边超过 1280 像素时，会从原图提取有内容的区域，按原始分辨率拼入一张局部识别图，
复用已加载的 YOLO 权重，最多追加一轮推理。整图检测保留，小控件候选回映到原图坐标并去重。
查看 `detail-input.png`、`detail-recovery.json` 中的区域映射、未容纳区域及裁剪边缘拒绝记录；
该步骤有一轮预算，不保证覆盖全部原始分辨率内容。MCP/Python 的 `detail_recovery=False`
或 CMD 的 `--no-detail-recovery` 可关闭；小图和模型图保持原来的单轮整图推理。
质量优先：默认保留原来的 `1280×1280` 方形 YOLO 输入。UI 模式的 YOLO 与 OCR 默认并行，共用 `--threads` 的线程预算：
默认 4 线程分配为 YOLO 3 线程、OCR 1 线程；单线程时串行运行。模型图沿用原来的方形、串行推理。
`raw-yolo.json` 保存实际输入形状，`inference_execution` 保存调度、线程分配和推理墙钟耗时；
并行时 `yolo_seconds` 与 `ocr_seconds` 存在重叠，不能直接相加作为总耗时。
减少填充属于实验选项：显式传 `--rectangular-inference` 或 MCP/Python 的 `rectangular_inference=True`，
才会按 32 像素步长减少方形填充和局部拼图末尾空白行。虽然检测缩放比例相同，模型输入形状变化仍可能改变预测框，
需要单独检查质量，不能视为无损提速。`--square-inference` 保留兼容，表示使用默认方形输入；
`--sequential-inference` / `parallel_inference=False` 用于串行对照。模型、阈值和补检轮次保持原配置。
空白候选必须有框内前景，或在 2 像素上下文中找到与候选匹配的完整闭合边框；仅有相邻面板线时会拒绝。
这样保留空输入框和复选框，同时阻止框外边线为纯白区域提供证据。`pixel_evidence` 记录
`inside_foreground_pixels`、`context_foreground_pixels`、`context_closed_border` 和 `pixel_support`。
其余与正常 OCR 文字重复的候选仍去重，不强制套用控制框图外框。识别结果仍可能漏检、误检或错字，
几何纠正只支持已实现的图形结构。

默认先自动裁剪可靠的窗口外框；`--no-auto-crop` 保留完整输入。`elements.json` 中
`crop.offset` 记录相对原图的偏移，标注和后续修正均对应 `analysis-input.png` 的尺寸。

两个图形配置还会检测彩色连通主体，不限定红、蓝、黄等颜色。相接的不同色块、被细黑线
分割的同一主体会合并；邻近文字和外部连线保持独立，邻接的闭合端口可纳入控件边界。
同时补充实心 Mux 条和文字触边的标签外框，过滤母线细节误认的接地符号。未匹配到已知
形状的 YOLO 候选，只要有足够的非文字图形像素，仍保留为待复核观察项。几何资源中记录
候选的采纳/过滤原因及颜色证据，汇总中记录颜色数量和融合耗时；未重新训练或替换模型。

默认自动补检已有闭合控件框内部遗漏的小文字：复用全图检测概率图，只在这些局部区域
检查较低分候选；要求原尺寸裁剪与放大裁剪的识别文字一致、至少一次分数达到 0.95，
两次均达到 0.90，并过滤重复文字、触及裁剪边缘的连线和边框。通过后追加文字项，保留
原有识别与编号，来源标记为 `rapidocr-small-text-recovery`。补检耗时、数量和证据保存在
`ocr-recovery.json`；可用 `--no-small-text-recovery` 关闭补检进行对照。

已有 YOLO 和 OCR 实验缓存时会自动复用。其他环境先在仓库目录安装依赖：

```cmd
python -m pip install ".[image]"
easy_uiauto_image "C:\images\input.bmp" "C:\images\results"
```

YOLO 权重首次下载约 281 MB，校验 SHA256 后缓存；后续复用。OCR 使用 RapidOCR 包
附带的 PP-OCRv6 ONNX 模型。`--offline` 要求已有缓存，禁止下载 YOLO 权重。
整个识别流程在本机 CPU 执行，不使用视觉 API，也不需要 API key。
仓库的 CMD 启动器使用 PATH 中的 `python`；如需指定解释器，可先设置
`EASY_UIAUTO_IMAGE_PYTHON`。用 `annotate-image.cmd --help` 查看全部参数。
成功返回退出码 0，失败返回 2；进入处理阶段后的失败记录在结果目录的 `summary.json`
和 `run.log` 中。

### 深度探索

深度探索会先执行轻度扫描，再操作已知且可恢复的命令，并记录操作造成的变化。
遇到新页面或弹窗时，可以在限制深度内继续学习。

深度学习的交互与录制回放共用 `Controller`，基础操作统一使用 `uiautomation`。
已校验的控件直接传给执行层，避免再次按 XPath 查找。操作默认显示同进程
`ScreenLineBox` 方框，并在获取前后证据截图前关闭；文本验证截图也会暂隐方框。
MCP 的 `learn_ui_command_effect` 和 `explore_ui_workflows` 可传入 `highlight=false`，
命令行对应 `--no-highlight`。人工录制仍独立保存动作。

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

### 执行已学习命令

```bash
easy_uiauto_ui search <app-id> "检索词"
easy_uiauto_ui show <app-id> --include executable
easy_uiauto_ui run <app-id> <页面.区域.控件.动作> --text "可选文本"
easy_uiauto_ui batch <app-id> '["main.keypad.enter-digit-6.click", "main.keypad.enter-addition-operation.click"]'
easy_uiauto_ui learn-effect <app-id> <command> --recover
easy_uiauto_ui teach <app-id> <control-id> "控件含义" intent "功能说明"
```

运行时依次使用唯一 `AutomationId` → 上下文 XPath → `Name`/`ClassName`/`ControlType`
属性组合 → 多状态控件图 → 本地 OCR → 显式开启的远程视觉。
最后一级需要传入 `--allow-vision-fallback` 或 `allow_vision_fallback=true`。
控件失效、缺失或匹配不唯一时会进入隔离区，不会继续点击。单条和批量命令默认会用
红框预览实际解析出的目标；可通过 `--no-highlight` 或 `highlight=false` 关闭约 100 ms
的可视预览等待。

`set-text` 会先验证 UIA 直接写入；失败时改用支持中文的剪贴板粘贴。只有可访问值
或从目标输入框精确复制回读确认文本已经写入，命令才会返回成功；控件局部画面变化
仅作为诊断信息。粘贴后会恢复原文本剪贴板。

批量执行只适用于同一个稳定页面，发生页面跳转后应拆成下一批。涉及外部影响或破坏性的
命令必须传入 `--confirm` 或 `confirm=true`。

### 知识库结构

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

扫描或操作响应学习提交记录后，会自动清理无引用的旧图片。每个控件保留当前模板和
最近 8 个不同的图片状态，超出范围且没有其他引用的状态图片会从磁盘删除。新交互截图
按 PNG 内容哈希去重，相同的操作前后图或桌面图共用一个文件。交互默认保留最近 50 次
操作的图片；图片过期后，文字交互记录仍然保留。

被知识记录引用的证据、标记了 `retain_images: true` 的交互，以及每个控件/动作最近一次
仍有图片的失败交互会额外保留。缺少图片清单的旧隔离控件也会保守保留证据。因此数量
上限不等于磁盘配额。可用 `EASY_UIAUTO_CONTROL_IMAGE_VARIANTS`（默认 `8`，最小 `1`）
和 `EASY_UIAUTO_INTERACTION_IMAGE_RECORDS`（默认 `50`，最小 `0`）调整默认保留数量。

已有应用可以调用 `cleanup_ui_knowledge_images(app_id="notepad")`，先查看预计删除的
文件和字节数，再传入 `dry_run=false` 实际清理。`keep_control_variants` 和
`keep_interaction_images` 可以覆盖本次调用的保留数量。清理前检查 Markdown/YAML 引用，
只删除该应用控件/交互目录中由程序管理的 PNG，不涉及页面图和其他应用。知识记录
无法读取时会推迟清理；知识库文件锁可防止多个新版 MCP 进程清掉彼此尚未提交的图片。

控件和区域只保存窗口内的 `normalized_rect`，并明确标记
`geometry_role: visual-hint-only`。它只用于页面标注和视觉对比，不能作为控件身份，
也不能直接用于点击。实时框选和操作会先重新解析当前 `LOCATION`，失败后才使用唯一
图片模板或本地 OCR 兜底。0.6 版重建旧知识库索引时，会自动清除历史绝对坐标、PID
和窗口句柄。

### 隐私与安全

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
`search_ui_knowledge`、`search_ui_knowledge_batch`、`list_ui_commands`、
`run_ui_command`、`run_ui_commands`、
`learn_ui_command_effect`、`explore_ui_workflows`、`list_ui_interactions`、
`teach_ui_control`、`rebuild_ui_knowledge_index` 和 `cleanup_ui_knowledge_images`。

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
`get_control_at_position` 的完整返回结果。旧的扁平参数仍兼容。存在 AutomationId 时会先
在目标窗口内查找，并忽略动态 Name、`foundIndex` 和历史坐标；ID 缺失或重复时再使用 XPath。

## Python API 示例

### 在代码中导入 MCP 工具

从 0.9.2 起，安装或升级 `easy-uiauto[mcp]>=0.9.2` 后，全部 54 个工具可从
`easy_uiauto.mcp` 导入。使用源码时，在编辑器选中的解释器中执行
`python -m pip install -e ".[mcp]"`。

```python
import json
from easy_uiauto.mcp import list_windows, get_control_tree, find_control

print(json.loads(list_windows()))
# 用上一步返回的实际窗口标题替换示例值。
tree = json.loads(get_control_tree("你的文件.txt - 记事本", max_nodes=500, timeout_seconds=15))
```

导入后直接在本进程中调用，不需要启动 MCP 服务。大部分工具返回 JSON 字符串，
`take_screenshot` 返回 MCP Image 对象。`scan_window_knowledge`、`explore_ui_workflows`、
`analyze_ui_image`、`correct_ui_image_annotations` 需要 `await`，普通 Python 可省略 `ctx`。
包根目录的传统 Controller 接口继续保留。运行 `python demo/mcp_import.py --list-tools`
可列出全部导出。完整目录、异步示例、返回类型和解释器排查见
[Python 导入说明](docs/python-imports.zh-CN.md)。

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

## 发布自动化

仓库通过 GitHub Actions 完成检查和发布：

- **CI**：使用 Python 3.11、3.12 和 3.13 运行测试，并检查构建结果。
- **Tag 发布**：通过 PyPI Trusted Publishing 构建并上传版本，无需维护 PyPI Token。
- **构建产物**：同时生成 sdist 和 wheel。

## 本地开发

```bash
pip install -e .[dev]
pytest
ruff check .
```

## 更多示例

更多示例请参考 `demo/` 目录中的测试文件或查看源代码中的文档字符串。

## 许可证

MIT License，详见 [LICENSE](https://github.com/Poggi-Tang/easyautomation/blob/main/LICENSE)。

## 联系方式

扫描以下二维码添加我的微信：

![微信二维码](https://github.com/Poggi-Tang/easyautomation/blob/main/src/image/or_code.bmp)
