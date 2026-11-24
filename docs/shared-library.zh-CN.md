# GUI、MCP 与 Python 共享知识库

三个入口使用同一份 `vault/applications/<app_id>` 数据。控件、区域、参考资料和组合动作的 Markdown 是原始记录；JSON 索引根据记录生成，过期后自动重建。Skill 只描述调用方法，不存放另一份控件数据。

## 选择同一个目录

使用 GUI 当前显示的知识库绝对路径。MCP 知识库接口、搜索、单控件命令、效果学习、整窗扫描和迁移接口均支持 `vault`。省略时使用配置的默认目录；GUI 临时切换目录不会改写其他进程的默认配置。

```python
import json
from easy_uiauto import KnowledgeLibrary
from easy_uiauto.mcp import get_ui_library, search_ui_knowledge

vault = r"D:\ui-knowledge"
library = KnowledgeLibrary(vault)
apps = library.apps()
app_id = apps[0]["id"]
controls = json.loads(get_ui_library(app_id, vault=vault, section="controls"))
matches = json.loads(search_ui_knowledge(app_id, query="新建", vault=vault))
```

`KnowledgeLibrary` 提供不依赖 MCP 服务运行的 Python 入口；`easy_uiauto.mcp` 下的函数也能直接导入调用，返回格式与 MCP 相同。已有区域学习和组合动作继续使用 `start_ui_library_task`，不需要启动 GUI。

## 采集与旧格式兼容

`capture_control_record_at_position`、`capture_control_record_on_ctrl` 只返回观察记录。使用 `upsert_control_record` 保存到共享知识库：

```python
from easy_uiauto.mcp import capture_control_record_at_position, upsert_control_record

record_json = capture_control_record_at_position(120, 240)
result = json.loads(upsert_control_record(
    record_json, app_id=app_id, vault=vault,
    page_id="manual-main", region_id="",  # 有所属区域时传实际完整区域 ID
))
```

纯 Python 等价调用为 `library.upsert_control_record(record, app_id=..., page_id=..., region_id=...)`，其中 `record` 是字典。同一软件、页面和定位的记录复用 ID；已有人工名称、说明、路径、图片和隔离状态保留。导入接口不用于改写已有控件定位。

不指定软件时，仅在窗口能唯一对应已保存软件时复用它；没有对应软件时可按记录中的软件信息创建。多个软件或页面状态存在歧义时返回错误，由调用方传入实际 ID。无区域信息的新记录存为未分区。

旧 `upsert_control_record` 名称保留，但已不再加载 `control_vector_store.py`，也不需要 `EASY_UIAUTO_CONTROL_VECTOR_DB_DIR`。没有截图、实时定位或操作结果证据的导入记录保持“已采集”，不能伪造为可执行或业务已验证。

## 迁移旧数据

本节适用于尚未入库的旧格式导出；修正已经入库的区域、控件和流程请看 [历史学习知识修复与转换指南](legacy-knowledge-repair.zh-CN.md)。

先从原存储导出 JSON 或 JSONL，再调用 `import_control_records`。接受记录数组、`{"records": [...]}`、单条记录或每行一条的 JSONL。支持旧 `LOCATION` 和小写 `xpath` 字段。不执行外部 Python 模块，不删除或修改原导出文件。

```python
from easy_uiauto.mcp import import_control_records

preview = json.loads(import_control_records(
    r"D:\backup\controls.json", app_id=app_id, vault=vault,
    page_id="manual-main", dry_run=True,
))
# 核对预览中的所属软件、ID 和失败行后执行。
result = json.loads(import_control_records(
    r"D:\backup\controls.json", app_id=app_id, vault=vault,
    page_id="manual-main", dry_run=False,
))
```

结果逐行报告成功或拒绝原因。成功行保留，修正失败行后可以重跑，相同控件不会重复生成。预览不写入控件记录，最终结果以实际导入返回为准。

## mixed 模式

`set_mode("mixed", vault=..., app_id=..., page_id=..., region_id=...)` 指定自动采集目的地。每次重新设置模式会更新本进程的目的地，`get_mode` 返回 `learning_context`。

动作返回的 `learning` 单独报告采集和入库结果。动作可能成功而入库失败，此时检查错误并修复归属或存储问题，不要为了入库而重复操作。mixed 是操作后的控件观察，不等同于深度探索或业务验证。

## 外部更新与正在编辑的内容

通过库接口写入后发布软件级变更标记。GUI 每 1.5 秒在后台检查标记，只读取有变化的软件并增量更新树，保留其他软件节点、展开状态和选择。GUI 正在运行任务或保存时延后刷新。

当前未保存的输入不会被刷新覆盖。如果同一记录已被其他入口修改或删除，GUI 拒绝覆盖并保留输入；可先复制编辑内容，重新选择最新记录后合并。关闭并重新打开的 GUI 会直接读取最新记录。直接用文本编辑器修改文件不属于这个通知协议，需要重新加载 GUI，检索前重建索引。

应用界面操作仍需串行执行；共享文件不意味着可以让多个 Agent 同时操作同一个软件窗口。
