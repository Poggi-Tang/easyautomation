# MCP 动作结果、等待与就绪检查

MCP 的桌面动作及定位工具返回 JSON 字符串，统一包含下列字段。Python
`Controller` 和 TCP 服务仍保留原接口；这里的约定针对 MCP。

```json
{
  "ok": true,
  "code": "action_executed",
  "message": "Clicked (left/single) at (100, 200)",
  "verified": false,
  "retryable": false,
  "data": {}
}
```

- `ok`：工具执行是否成功；不等于目标业务流程完成。
- `code`：供程序判断的稳定结果代码。
- `verified`：是否有读取结果或条件检查的证据。定位成功只验证定位；普通点击不验证业务结果。
- `retryable`：是否明确允许重试当前工具。默认 `false`；不代表重新点击一定安全。
- `message`：可读说明；旧文本返回值迁移到这里。
- `data`：原始结构化结果或等待的最后一次观测。原有 JSON 业务字段同时保留在顶层。

| code | 含义 | 建议处理 |
| --- | --- | --- |
| `action_executed` | 动作已执行，业务结果未验证 | 检查预期文本、控件或窗口条件 |
| `verified` | 指定观察条件已成立 | 根据被验证的条件继续流程 |
| `control_not_found` | 没找到控件 | 检查定位器，必要时等待出现 |
| `ambiguous_match` | 匹配不唯一 | 补充 XPath、属性或明确的 `found_index` |
| `window_invalid` | 窗口不存在、失效或无法访问/激活 | 重新获取窗口，不向其他窗口发送输入 |
| `control_disabled` | 控件当前不可用 | 检查前置条件，等待 `enabled` |
| `action_failed` | 动作或底层观察失败 | 读取错误和当前状态，可能已有部分操作发生 |
| `result_unconfirmed` | 指定结果未得到确认或等待超时 | 先观察当前状态，不自动重复原动作 |
| `invalid_argument` | 参数无效 | 修改参数 |
| `operation_blocked` | 当前模式或确认要求禁止操作 | 检查模式和该操作的确认要求 |

`run_actions` 中每一步包含相同字段；`stop_on_error=true` 能识别 JSON
`ok=false`、`found=false` 和原有文本错误。批量失败时保留已经执行的前缀，不能直接重跑整批。

## 等待和验证

新增工具只观察，不点击、不激活窗口，也不自动重放操作。超时范围为 0–120 秒；
`timeout_seconds=0` 只检查一次。轮询使用单调时钟；正在执行的原生 UIA 调用无法强制中断，
因此可能超过轮询期限，超时观察不会被当作按时成功。

以下示例表示 MCP 工具调用参数：

```python
location = {
    "WindowName": "Editor",
    "AutomationId": "status",
    "ControlType": "TextControl"
}

wait_for_control(location=location, state="exists", timeout_seconds=10)
wait_for_control(location=location, state="absent", timeout_seconds=5)
wait_for_control(location=location, state="enabled", timeout_seconds=5)
wait_for_text(location=location, expected_text="Saved", match="exact", timeout_seconds=5)
wait_for_text(location=location, expected_text="Saved", match="contains", timeout_seconds=0)
wait_for_text(location=location, match="changed", baseline_text="Saving", timeout_seconds=5)
```

文本来自 UIA Value/LegacyIAccessible/Text pattern，不使用剪贴板读取，也不把控件名称
自动当作正文。无法读取文本时不会报告验证成功。`changed` 必须提供动作发生前的基准文本。

```python
before = get_window_state(window_name="Editor")
# 解析返回 JSON，保存 before["data"]["observation"]，然后执行目标动作。
wait_for_window(
    window_name="Editor", state="changed",
    baseline=before["data"]["observation"], timeout_seconds=5
)
wait_for_window(window_name="Confirmation", state="exists", timeout_seconds=5)
wait_for_window(window_name="Confirmation", state="absent", timeout_seconds=5)
```

窗口按精确标题匹配；重复标题返回 `ambiguous_match`。`changed` 仅比较提供的窗口属性
（标题、句柄、进程、边界、可用/可见状态等），不代表窗口内部任意内容变化。基准带有句柄时，
跟踪同一窗口以识别标题变化。窗口关闭需用 `absent` 明确验证。句柄和 PID 只用于当前会话观测。

控件查找默认要求唯一匹配；明确的正数 `found_index` 为 1 起算的选择。
等待控件消失时，父窗口失效、匹配不唯一或 UIA 异常均不会被当作控件消失。

## 滚动工具

新流程统一使用 `scroll(amount=3, direction="down", x=-1, y=-1)`。
`amount` 必须是正整数；`direction` 只能是 `up` 或 `down`，分别发送正、负滚轮刻度。
`mouse_scroll` 保留为兼容别名，参数和实现与 `scroll` 相同。

默认 `x=y=-1` 使用当前鼠标位置；指定位置时同时提供两项，支持副屏负坐标，
但 `-1` 仍保留为省略坐标的哨兵值。旧 `scroll` 上下方向反转的问题已修正。

## 可选能力就绪状态

`get_capability_readiness(language="eng+chi_sim")` 分别报告：

- `uia`：Windows、UIA 模块和桌面根节点访问。
- `ocr`：Python 适配器、Tesseract 可执行文件和请求的语言数据。
- `image_matching`：OpenCV、NumPy、Pillow 和合成图片匹配检查。
- `remote_vision`：当前配置是否齐全、URL 格式是否有效。

每项包含 `configured`、`available`、`runtime_verified`、`missing`、`issues` 和
`scope`。按 `scope` 理解可用范围：本地引擎通过不代表桌面截图或所有 UI 操作可用。
远程视觉仅配置齐全时是 `status="unverified"`、`available=null`；没有调用网络进行验证。

诊断不会安装依赖、截图、上传数据或返回密钥。本地子进程检查有超时限制。
原有 `get_ui_learning_readiness` 保留，用于学习配置和知识库状态；它不替代分项能力检查。

## 验证范围

新增跨平台单元测试使用虚拟时钟和模拟 UIA/输入回调，验证错误分类、停止条件、参数校验和
工具注册；它们不证明真实 Windows 桌面操作成功。完整项目测试和桌面实测仍需 Windows。
