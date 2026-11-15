"""YOLO device selection. Import torch only inside an inference worker."""
from .image_model_paths import yolo_device


def inspect_device(device=""):
    requested = yolo_device(device)
    import torch

    report = dict(requested=requested, device="cpu", cuda_available=False,
                  torch_version=str(torch.__version__), cuda_version=torch.version.cuda,
                  gpu_name="", fallback_reason="", ok=True)
    if requested == "cpu":
        return report  # Explicit CPU must not initialize a GPU or require its driver.
    selected = "cuda:0" if requested in {"auto", "cuda"} else requested
    reason = ""
    if not torch.version.cuda:
        reason = "当前运行环境是 CPU 版 PyTorch；需要安装 NVIDIA GPU 运行环境"
    else:
        try:
            if not torch.cuda.is_available():
                reason = "CUDA 不可用；请检查 NVIDIA 显卡和驱动"
            elif int(selected.split(":")[1]) >= torch.cuda.device_count():
                reason = "指定的 NVIDIA GPU 编号不存在"
            else:
                # Availability alone misses incompatible driver/architecture builds.
                probe = torch.ones(1, device=selected) + 1
                torch.cuda.synchronize(selected)
                del probe
                report.update(cuda_available=True, device=selected,
                              gpu_name=torch.cuda.get_device_name(selected))
        except (RuntimeError, OSError, AssertionError) as error:
            reason = f"CUDA 初始化失败（{type(error).__name__}）；请检查驱动与 PyTorch 的兼容性"
    if reason:
        report.update(ok=requested == "auto", device="cpu" if requested == "auto" else "",
                      fallback_reason=reason)
    return report


def require_device(device=""):
    report = inspect_device(device)
    if not report["ok"]:
        raise RuntimeError(report["fallback_reason"] + "；可选择“自动”或 CPU")
    return report


if __name__ == "__main__":
    import json
    import sys
    print(json.dumps(inspect_device(sys.argv[1] if len(sys.argv) > 1 else ""), ensure_ascii=False))
