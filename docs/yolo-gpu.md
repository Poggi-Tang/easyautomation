# YOLO GPU setup

English | [简体中文](yolo-gpu.zh-CN.md)

YOLO supports CPU and NVIDIA CUDA. OCR stays on CPU. A compatible NVIDIA GPU and
driver are required; a separate CUDA Toolkit is unnecessary. Intel / AMD GPU inference
is not supported. This feature is in the current source; published 1.1.0 bundles
predate it. Install from the updated source or use a subsequent release.

## Install

Run `setup.bat -Gpu` from the repository. It prepares the GUI and CPU environment,
then downloads optional CUDA packages from the official PyTorch index. Allow several
GB of network traffic and disk space. The default index is `cu128`; for another
compatible official index use, for example, `-CudaIndex cu130`. It must provide a
wheel for your Python / Windows version and support your GPU and driver.

CUDA packages live in a separate directory. They become active only after CUDA
initialization and real YOLO inference succeed. Failure preserves the previous
runtime. Restart GUI / MCP after installation. Future complete GUI ZIPs support
the same command. Ordinary `setup.bat` installations continue to include CPU packages.

For installations made with `install-full.cmd` or pip, run from the updated checkout:

```bat
"C:\path\to\python.exe" -I -X utf8 scripts\install_gpu.py
```

Use the Python that runs MCP, with the updated source, image dependencies and verified
YOLO weights already installed. GUI ZIPs use `_internal\image-runtime\python.exe`;
source GUI setup defaults to `%LOCALAPPDATA%\easy-uiauto\gui\python\python.exe`.
Each installation directory needs its own CUDA environment.

## Select and inspect

- GUI: local annotation models → YOLO inference device → Auto / CPU / NVIDIA GPU.
- CLI: `easy_uiauto_image input.png output --device cuda`.
- MCP: `analyze_ui_image(image_path="input.png", device="cuda")`, with `engine="local"`.
- Python: `analyze_local_image("input.png", "output", device="cuda:0")`, imported from `easy_uiauto.image_cli`.

Precedence: call argument → `EASY_UIAUTO_YOLO_DEVICE` environment variable → saved
GUI preference → `auto`. Preferences share `%LOCALAPPDATA%\easy-uiauto\model-paths.json`
with model paths, and apply to GUI, MCP and Python under the same user.

`auto` prefers the first NVIDIA GPU and falls back to CPU on CUDA initialization
or inference failure. `cpu` forces CPU. `cuda` / `cuda:N` require the selected GPU
(zero-based index) and fail explicitly. Selecting a GPU does not install packages.
Check `device`, `requested_device` and `fallback_reason` in `raw-yolo.json`, or
`inference_execution.yolo_device` in the summary, for actual execution.

GPU acceleration applies to YOLO forward inference. OCR, postprocessing, model
loading and screenshots still take time; total speedup depends on the workload.

## Return to CPU

Select CPU in the GUI, or disable the optional packages with the appropriate runtime:

```bat
"C:\path\to\python.exe" -I -X utf8 scripts\install_gpu.py --disable
```

Restart and select Auto or CPU. Disabling keeps downloaded CUDA files; repeated
installations also retain previous successful directories. Moving the complete GUI
folder preserves the active CUDA environment.
