# DeepSeek Harness（DSH）配置

通过 DSH 官方的 `@deepseek-ai/dsh-mcp-client` 插件，以 stdio 方式启动
easy-uiauto。下面使用 DSH CLI / Web 的 `--patch` 配置方式。
官方配置格式参见 [MCP 客户端](https://github.com/deepseek-ai/deepseek-harness/blob/master/packages/mcp/mcp-client/README.md)
和 [MCP overlay 示例](https://github.com/deepseek-ai/deepseek-harness/blob/master/docs/user/guide/mcp-memory.md)。

## 1. 准备 Windows Python 环境

DSH Host、easy-uiauto 和目标软件应运行在同一台 Windows 电脑的交互式桌面会话中。
远程 Linux Host 或 WSL 中启动的 Python 不能直接操作 Windows 本机的软件界面。

已有 Python 3.11 或更高版本时，在 CMD 或 PowerShell 中安装：

```cmd
python -m pip install -U "easy-uiauto[mcp,vision]" -i https://pypi.org/simple
python -c "import sys, easy_uiauto; print(sys.executable); print(easy_uiauto.__version__)"
```

记下输出的 `python.exe` 完整路径，后续 MCP 和 Skill 安装均使用这个解释器。
`mcp,vision` 包含 MCP、本地图片匹配和 OCR 的 Python 依赖；OCR 还需要 Tesseract
可执行文件及语言包。需要 YOLO＋RapidOCR 图片标注及完整运行环境时，使用
[README 中的一键完整安装](../README.zh-CN.md#一键完整安装与离线部署)。
完整安装默认解释器为 `%LOCALAPPDATA%\easy-uiauto\full\python\python.exe`；
以实际安装目录为准。

## 2. 配置知识库和视觉接口

在上述 Python 环境中运行：

```cmd
easy_uiauto_config
```

填写知识库路径、API URL、API Key 和 API Model，保存即可。若命令不在 PATH 中，
使用该环境 `Scripts` 目录中的 `easy_uiauto_config.exe`；完整安装也提供
`easy_uiauto_config.cmd`。

DSH 对话模型和 easy-uiauto 视觉接口分别配置：DSH 的模型负责调用工具，
easy-uiauto 的界面扫描、语义分析和交互效果分析使用这里保存的视觉接口。
视觉模型须支持 OpenAI 兼容的图片输入；仅配置 DSH 的模型和 Key 不会完成这一步。
基础 UIA 操作不要求视觉 API，本地 YOLO＋RapidOCR 图片标注也不使用这个 API。

## 3. 添加 MCP

下载或复制 [easy-uiauto.cordis.yml](examples/dsh/easy-uiauto.cordis.yml)，
保存为例如 `C:\configs\easy-uiauto.cordis.yml`：

```yaml
- insert:
    - id: mcp-easy-uiauto
      name: '@deepseek-ai/dsh-mcp-client'
      config:
        serverName: easy_uiauto
        transport: stdio
        command: 'C:/Python314/python.exe'
        args: ['-u', '-m', 'easy_uiauto.mcp.server']
        toolCallTimeoutMs: 600000
        env:
          PYTHONIOENCODING: utf-8
          PYTHONUNBUFFERED: '1'
```

将 `command` 替换为第 1 步输出的解释器路径。Windows 路径可以使用正斜杠，
路径中有空格时也保持为一个字符串；不要在路径内容中再加入双引号。
`args` 已单独传入，无需添加 `cmd /c`。该示例用于 DSH 原生 overlay，
不使用其他客户端的 `mcpServers` JSON 格式。

通过已有 DSH CLI 启动：

```cmd
dsh web --patch "C:\configs\easy-uiauto.cordis.yml"
```

没有全局 `dsh` 命令时，可使用官方 npm 启动方式：

```cmd
npx @deepseek-ai/dsh web --patch "C:\configs\easy-uiauto.cordis.yml"
```

当前 DSH CLI 的 profile 配置位于 `$DSH_HOME/profiles/<profile>/cordis.patch.yml`；
未设置 `DSH_HOME` 时，Web profile 默认位于
`%USERPROFILE%\.dsh\profiles\web\cordis.patch.yml`。
如需持久启用，把示例的 `insert` 项合并到该文件的顶层 YAML 列表中，
之后使用 `dsh web` 即可；不要覆盖已有配置，也不要同时通过 `--patch` 重复加载。
详见 [DSH profile 配置](https://github.com/deepseek-ai/deepseek-harness/blob/master/apps/cli/README.md#profiles)。

DSH 会管理 stdio 子进程，无需再单独打开一个 `easy_uiauto` 服务窗口。
这里连接的是 MCP 服务，不是 `easy_uiauto_service` TCP 服务。

## 4. 安装配套 Skill

将随包提供的 `easy-uiauto-learning` 和 `easy-uiauto-operate` 两个目录复制到
`$DSH_HOME/skills`，默认是 `%USERPROFILE%\.dsh\skills`。DSH 通过目录中的
`SKILL.md` 发现技能；只注册 MCP 不会自动复制这些文件。
参见 [DSH 本地 Skill 目录说明](https://github.com/deepseek-ai/deepseek-harness/blob/master/packages/skill/skill-filesystem/README.md#roots-and-priority)。

在 PowerShell 中执行，`$uiaPython` 使用第 1 步的解释器路径。
此命令会安装或更新这两个同名 Skill：

```powershell
$uiaPython = 'C:\Python314\python.exe'
$skillSource = & $uiaPython -c "from easy_uiauto.mcp.skill_installation import bundled_skills_dir; print(bundled_skills_dir())"
if ($LASTEXITCODE -ne 0) { throw '无法读取 easy-uiauto Skill 目录，请检查解释器和安装环境' }
$harnessDir = if ($env:DSH_HOME) { $env:DSH_HOME } else { Join-Path $env:USERPROFILE '.dsh' }
$skillTarget = Join-Path $harnessDir 'skills'
New-Item -ItemType Directory -Path $skillTarget -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $skillSource 'easy-uiauto-learning') -Destination $skillTarget -Recurse -Force
Copy-Item -LiteralPath (Join-Path $skillSource 'easy-uiauto-operate') -Destination $skillTarget -Recurse -Force
```

目录结构应为：

```text
<DSH_HOME>/skills/
├── easy-uiauto-learning/SKILL.md
└── easy-uiauto-operate/SKILL.md
```

在 DSH 中确认两个 Skill 已出现在技能列表中，并要求模型读取对应 Skill 后执行任务。
若当前会话未更新技能列表，重新进入会话或重启 DSH。
本项目现有的 `--install-codex-skills` 只安装到 Codex，不能代替这一步。

## 5. 验证与使用

连接成功后，应出现 `mcp__easy_uiauto__list_windows` 等工具。
先在 DSH 中输入：

```text
调用 easy_uiauto 的 get_capability_readiness 和 get_ui_learning_readiness，
再调用 list_windows。报告知识库路径、各项能力状态和窗口列表，不执行点击。
```

`get_ui_learning_readiness` 检查学习配置是否齐全，不会联网验证 API 或模型是否可用。
知识库路径以该工具实际返回的 `knowledge_vault` 为准。

基础扫描示例：

```text
读取 easy-uiauto-learning 的 SKILL.md，学习当前打开的记事本界面。
使用 visual-first 扫描并显示标注，报告可执行命令和未验证项。
```

深度探索示例：

```text
读取 easy-uiauto-learning 的 SKILL.md，深度学习当前打开的记事本。
使用 safe 策略，最多操作 10 次，递归深度 3，保存操作前后的证据和恢复结果。
```

使用已有知识库时，读取 `easy-uiauto-operate`，先搜索并核验对应命令，
再通过 `run_ui_command` 操作。

## 常见问题

- **模块不存在或工具未出现**：检查 `command` 是否指向安装了 MCP 扩展的 Python，
  用该解释器运行 `-m easy_uiauto.mcp.server --version`，再查看 DSH 的 MCP 启动日志。
- **学习等待超时**：示例的 `toolCallTimeoutMs: 600000` 表示单次等待最多 10 分钟，
  不是任务固定耗时。超时后先调用 `get_ui_learning_status(task_id=...)`；未取得任务 ID
  时可省略参数查询最近任务。仍在运行就继续查询，不要重复扫描或杀掉 MCP 进程。
- **视觉接口未配置**：检查配置窗口和 `get_ui_learning_readiness`，不要仅修改 DSH
  对话模型。DSH 会过滤子进程的部分敏感环境变量；如改用进程环境传递 Key，需在 MCP
  的 `config.env` 中显式传入 `EASY_UIAUTO_VISION_API_KEY`。
  参见 [DSH stdio 环境说明](https://github.com/deepseek-ai/deepseek-harness/blob/master/packages/mcp/mcp-client/src/transport.ts)。
- **找不到 UIA3 或要求切换 UIA2**：本项目使用 Python `uiautomation`，无需另装名为
  UIA3 的后端。检查实际工具错误和扫描 `failures`，不要根据视觉绑定失败推断 Qt 没有暴露控件。
- **扫描完成但命令为 0**：检查每项失败原因和隔离状态。`ok=true` 仅表示扫描完成，
  截图和视觉标注本身不代表已经得到可执行 XPath。
