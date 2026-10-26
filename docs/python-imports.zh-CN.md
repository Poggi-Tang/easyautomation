# 在 Python 中导入 MCP 工具

从 0.9.2 起，所有 54 个 MCP 工具都可以从 `easy_uiauto.mcp` 导入。导入后是在当前 Python
进程中直接调用函数，不会启动 MCP 服务，也不会连接 Codex 已经启动的 MCP 进程。

## 安装与解释器

在写代码所用的解释器中安装，安装名是 `easy-uiauto`，导入名是 `easy_uiauto`：

```cmd
python -m pip install --upgrade "easy-uiauto[mcp]>=0.9.2"
```

使用本仓库源码时，在仓库根目录执行：

```cmd
.venv\Scripts\python.exe -m pip install -e ".[mcp,vision]"
```

然后在编辑器中选择该 `.venv\Scripts\python.exe`。系统 Python 和项目虚拟环境可能装了
不同版本；不要只看终端当前目录。检查实际使用的解释器、版本和模块位置：

```cmd
python -c "import sys,easy_uiauto; print(sys.executable); print(easy_uiauto.__version__); print(easy_uiauto.__file__)"
```

如果 `from easy_uiauto.mcp import get_control_tree` 报 `ImportError`，先核对上述路径，
确认它指向当前源码或包含这些导出的新版本。源码新增了显式类型导入、`__all__`、
`__dir__` 和 `py.typed`，用于编辑器补全和类型信息。

## 同步工具

```python
import json
from easy_uiauto.mcp import get_control_tree, list_windows, find_control, wait_for_control

windows = json.loads(list_windows())
print([window["title"] for window in windows])

# 填入上一步实际返回的完整窗口标题。
window_name = "你的文件.txt - 记事本"
tree = json.loads(get_control_tree(window_name, max_depth=5, max_nodes=500, timeout_seconds=15))
print(tree)

# 控件属性来自该窗口的实际树/已记录 LOCATION，不猜测 AutomationId。
location = {"WindowName": window_name, "AutomationId": "15", "ControlType": "EditControl"}
found = json.loads(find_control(location=location))
if not found["ok"]:
    raise RuntimeError(found["message"])
ready = json.loads(wait_for_control(location, state="enabled", timeout_seconds=3))
print(ready)
```

上面的 `15` 是经典记事本编辑区的示例值，其他版本或软件必须使用实测属性。
文件修改、页面跳转等操作可能改变窗口标题（例如前面增加 `*`），后续调用应重新读取
实际标题。后台窗口受 Windows 前台限制时会返回 `window_invalid`，不能据此继续向其他
窗口输入。`run_action` 回放使用采集或录制得到的完整 `LOCATION`，不能只传入其中几个
选择器字段；`find_control` 等定位工具则支持规范化的简短 LOCATION。
`get_control_tree` 默认最多 500 个节点、15 秒，包含独立进程的启动时间。正常结果是嵌套
树，`_inspection` 标明达到的上限或提供程序错误；超时是 `ok=false`、`code="timeout"`。

可直接运行仓库的只读示例：

```cmd
.venv\Scripts\python.exe demo\mcp_import.py
.venv\Scripts\python.exe demo\mcp_import.py --window "你的文件.txt - 记事本"
.venv\Scripts\python.exe demo\mcp_import.py --list-tools
```

## 异步工具

以下四个函数需要 `await`：`scan_window_knowledge`、`explore_ui_workflows`、
`analyze_ui_image`、`correct_ui_image_annotations`。普通 Python 可以省略 `ctx`，
MCP 客户端调用时仍由 SDK 注入进度上下文，`ctx` 不会出现在工具参数 schema 中。

```python
import asyncio
import json
from easy_uiauto.mcp import analyze_ui_image, correct_ui_image_annotations

async def main():
    raw = await analyze_ui_image(image_path=r"C:\images\ui.png", output_dir=r"C:\images\results")
    analysis = json.loads(raw)
    if analysis.get("ok") is False:
        raise RuntimeError(analysis.get("error", analysis))
    corrected = await correct_ui_image_annotations(
        coordinates_json=analysis["coordinates_json"], profile="icons"
    )
    print(json.loads(corrected))

asyncio.run(main())
```

`analyze_ui_image` 会把指定图片发给已配置的视觉 API；修正工具只在本地处理。
扫描、响应学习和工作流探索也需要视觉配置。已处于异步环境时直接 `await main()`，
不要在运行中的事件循环里再调用 `asyncio.run()`。

## 返回类型与运行状态

- 大多数工具返回 JSON **字符串**，使用 `json.loads()`。部分旧错误路径返回以
  `Error:` 开头的纯文本；无法解析 JSON 时应保留并报告原始错误。
- `list_windows` 等列表查询解析后是列表；动作工具通常返回包含 `ok`、`code`、
  `verified`、`message` 的对象。执行成功不等于业务效果已验证，需要读取或等待结果。
- `take_screenshot` 在 Python 中返回 MCP `Image` 对象，使用其 `data` 保存 PNG：

```python
from pathlib import Path
from easy_uiauto.mcp import take_screenshot

image = take_screenshot(region_x=0, region_y=0, region_width=400, region_height=300)
Path("screen.png").write_bytes(image.data)
```

- `set_mode`、学习任务状态保存在当前进程中。Python 导入调用与 Codex 的常驻 MCP
  不共享内存状态。`EASY_UIAUTO_KNOWLEDGE_DIR` 可指定相同的文件知识库；进程间图片
  写入与清理使用知识库锁。
- 包根目录的 `Controller`、`find_control` 等保留传统接口。MCP 的同名函数请从
  `easy_uiauto.mcp` 导入；它们有自己的参数、结果和验证约定。
- `from easy_uiauto.mcp.server import get_control_tree` 也可用；公开入口推荐使用
  `easy_uiauto.mcp`。不要调用 `mcp.run()` 来完成普通函数调用。

## 完整导入目录

下表中的每一个名称都支持 `from easy_uiauto.mcp import 名称`。额外依赖仍按功能安装；
图像匹配/OCR 使用 `vision` 扩展，OCR 还需要 Tesseract。向量写入需要另外配置
`EASY_UIAUTO_CONTROL_VECTOR_DB_DIR` 指向真实的 `control_vector_store.py` 集成目录。

| 功能 | 可导入函数 |
| --- | --- |
| 模式 | `get_mode`, `set_mode` |
| 窗口 | `list_windows`, `activate_window`, `get_window_state`, `wait_for_window` |
| 控件与等待 | `get_control_tree`, `find_control`, `get_control_at_position`, `wait_for_control`, `wait_for_text` |
| 采集与向量记录 | `capture_control_record_at_position`, `capture_control_record_on_ctrl`, `upsert_control_record` |
| 就绪检查 | `get_capability_readiness`, `get_ui_learning_readiness` |
| 学习与探索 | `scan_window_knowledge`, `get_ui_learning_status`, `learn_ui_command_effect`, `explore_ui_workflows` |
| 知识检索 | `list_ui_knowledge_apps`, `search_ui_knowledge`, `search_ui_knowledge_batch`, `list_ui_commands`, `list_ui_interactions`, `show_ui_controls` |
| 知识操作与维护 | `run_ui_command`, `run_ui_commands`, `teach_ui_control`, `revalidate_ui_control`, `rebuild_ui_knowledge_index`, `cleanup_ui_knowledge_images` |
| 鼠标 | `click`, `click_at_position`, `move_mouse`, `drag_control`, `scroll`, `mouse_scroll` |
| 键盘 | `type_text`, `set_text`, `press_key`, `hotkey` |
| 批量动作 | `run_action`, `run_actions` |
| 截图与分割 | `take_screenshot`, `detect_visual_elements` |
| 图片分析与修正 | `analyze_ui_image`, `correct_ui_image_annotations` |
| 模板、OCR 与视觉定位 | `find_control_by_image`, `click_by_image`, `find_text_on_screen`, `click_text_on_screen`, `find_control_by_vision`, `click_by_vision` |
