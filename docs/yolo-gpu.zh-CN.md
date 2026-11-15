# YOLO GPU 使用说明

[English](yolo-gpu.md) | 简体中文

YOLO 支持 CPU 和 NVIDIA CUDA；OCR 仍使用 CPU。需要兼容的 NVIDIA 显卡和驱动，
无需单独安装 CUDA Toolkit。Intel / AMD 显卡暂不支持 GPU 推理。
本功能从 1.1.1 起提供；使用 1.1.1 或更新版本的源码、GUI 包或 Python 包。

## 安装

在仓库目录打开终端，执行：

```bat
setup.bat -Gpu
```

此命令先准备 GUI 和 CPU 环境，再从 PyTorch 官方源下载可选 CUDA 依赖，可能需要数 GB
下载和磁盘空间。默认使用 CUDA 12.8 轮子源，可用 `-CudaIndex cu130` 指定其他官方源；
该源必须提供当前 Python / Windows 对应的轮子，并与显卡和驱动兼容。
安装期间不替换原 CPU 包，CUDA 初始化和一次真实 YOLO 推理成功后才启用。
失败不会启用新环境，修复驱动后可重试。安装后重启 GUI / MCP。

1.1.1 起的完整 GUI ZIP 也使用相同命令。默认双击 `setup.bat` 仍安装 CPU 环境。
通过 `install-full.cmd` 或 pip 安装的用户，在当前源码仓库中用**运行 MCP 的 Python**执行：

```bat
"C:\path\to\python.exe" -I -X utf8 scripts\install_gpu.py
```

该 Python 中需要安装当前源码及本地图片依赖，并已有校验通过的 YOLO 权重。
GUI ZIP 的推理 Python 位于 `_internal\image-runtime\python.exe`；源码 GUI 的默认位置为
`%LOCALAPPDATA%\easy-uiauto\gui\python\python.exe`。不同安装目录各自准备 GPU 环境。

## 选择与检查设备

- GUI：“本地标注模型” → “YOLO 推理设备” → 自动 / CPU / NVIDIA GPU。
- CLI：`easy_uiauto_image input.png output --device cuda`。
- MCP：`analyze_ui_image(image_path="input.png", device="cuda")`，限 `engine="local"`。
- Python：`analyze_local_image("input.png", "output", device="cuda:0")`，从 `easy_uiauto.image_cli` 导入。

优先级为：调用参数 → `EASY_UIAUTO_YOLO_DEVICE` 环境变量 → GUI 保存的选择 → `auto`。
GUI 配置保存在当前用户的 `%LOCALAPPDATA%\easy-uiauto\model-paths.json`，与模型路径共存；
GUI、MCP 和 Python 在同一用户下读取相同配置。

`auto` 优先使用第一块可用的 NVIDIA GPU，初始化或推理失败时使用 CPU，并记录原因。
`cpu` 强制使用 CPU；`cuda` / `cuda:N` 强制使用指定 GPU（编号从 0 开始），失败直接报错。
选择 GPU 不会自动下载依赖。`raw-yolo.json` 中的 `device` 表示实际设备，
`requested_device` 表示请求模式，`fallback_reason` 表示回退原因；汇总中的
`inference_execution.yolo_device` 也会记录实际设备。

GPU 加速 YOLO 的前向推理。OCR、部分后处理、模型加载和截图仍有开销，整体耗时不会按固定倍数缩短。

## 恢复 CPU 环境

可以直接在 GUI 中选择 CPU。需要停用可选 CUDA 包时，用相应推理 Python 执行：

```bat
"C:\path\to\python.exe" -I -X utf8 scripts\install_gpu.py --disable
```

重启后加载原 CPU 包，并将设备改为自动或 CPU。停用保留下载好的 GPU 目录，重复安装也保留
旧的成功安装目录；整个 GUI 文件夹移动后仍可使用已启用的 GPU 环境。
