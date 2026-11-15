"""Device selection and failures without requiring a CUDA machine."""
import sys
from types import SimpleNamespace

import pytest

from easy_uiauto import image_detection, image_model_paths, inference_device


@pytest.fixture(autouse=True)
def isolated_config(tmp_path, monkeypatch):
    monkeypatch.setenv(image_model_paths.CONFIG_ENV, str(tmp_path / "models.json"))
    monkeypatch.delenv(image_model_paths.DEVICE_ENV, raising=False)


@pytest.fixture
def torch_stub(monkeypatch):
    calls = []
    torch = SimpleNamespace(__version__="test", version=SimpleNamespace(cuda="12.8"),
        cuda=SimpleNamespace(is_available=lambda: True, device_count=lambda: 2,
            get_device_name=lambda device: "Test GPU " + device,
            synchronize=lambda device: calls.append(("sync", device)),
            empty_cache=lambda: calls.append(("empty",))),
        ones=lambda n, *, device: calls.append(("allocate", device)) or 1)
    monkeypatch.setitem(sys.modules, "torch", torch)
    return torch, calls


def test_explicit_cpu_does_not_initialize_cuda(torch_stub):
    torch, calls = torch_stub
    torch.cuda.is_available = lambda: pytest.fail("CPU must not initialize CUDA")
    report = inference_device.require_device("cpu")
    assert report["device"] == "cpu" and not calls and not report["fallback_reason"]


@pytest.mark.parametrize("requested,actual", [("auto", "cuda:0"), ("cuda", "cuda:0"), ("cuda:1", "cuda:1")])
def test_gpu_selection_allocates_on_the_requested_device(torch_stub, requested, actual):
    _, calls = torch_stub
    report = inference_device.require_device(requested)
    assert report["device"] == actual and report["cuda_available"] and report["gpu_name"]
    assert calls == [("allocate", actual), ("sync", actual)]


@pytest.mark.parametrize("failure", ["cpu_build", "unavailable", "driver", "no_devices"])
def test_auto_falls_back_but_explicit_cuda_fails(torch_stub, failure):
    torch, _ = torch_stub
    if failure == "cpu_build":
        torch.version.cuda = None
    elif failure == "unavailable":
        torch.cuda.is_available = lambda: False
    elif failure == "driver":
        def broken(*args, **kwargs):
            raise RuntimeError("unsupported driver")
        torch.ones = broken
    else:
        torch.cuda.device_count = lambda: 0
    report = inference_device.require_device("auto")
    assert report["device"] == "cpu" and report["fallback_reason"]
    with pytest.raises(RuntimeError):
        inference_device.require_device("cuda")


def test_invalid_gpu_index_reports_error(torch_stub):
    with pytest.raises(RuntimeError, match="编号"):
        inference_device.require_device("cuda:2")


def test_runtime_failure_retries_entire_yolo_job_only_in_auto(torch_stub, monkeypatch):
    _, calls = torch_stub
    attempts = []
    def infer(image, weights, **options):
        attempts.append(options)
        if options["device"].startswith("cuda"):
            raise RuntimeError("out of memory during detail recovery")
        return dict(device="cpu", boxes=[[1, 2, 3, 4]])
    monkeypatch.setattr(image_detection, "_run_yolo_on_device", infer)
    result = image_detection.run_yolo(None, "unused.pt", device="auto", recover_details=True)
    assert [a["device"] for a in attempts] == ["cuda:0", "cpu"]
    assert all(a["recover_details"] for a in attempts)
    assert result["device"] == "cpu" and result["requested_device"] == "auto"
    assert "out of memory" in result["fallback_reason"] and ("empty",) in calls
    attempts.clear()
    with pytest.raises(RuntimeError, match="out of memory"):
        image_detection.run_yolo(None, "unused.pt", device="cuda")
    assert len(attempts) == 1


def test_cpu_failure_is_not_retried(torch_stub, monkeypatch):
    attempts = []
    def broken(*args, **options):
        attempts.append(options)
        raise RuntimeError("invalid model")
    monkeypatch.setattr(image_detection, "_run_yolo_on_device", broken)
    with pytest.raises(RuntimeError, match="invalid model"):
        image_detection.run_yolo(None, "unused.pt", device="cpu")
    assert len(attempts) == 1


@pytest.mark.parametrize("device", ["cpu", "cuda:0"])
def test_tensor_decoding_preserves_original_coordinates(device):
    torch = pytest.importorskip("torch")
    from PIL import Image
    if device.startswith("cuda") and not torch.cuda.is_available():
        pytest.skip("CUDA hardware unavailable")
    class Model:
        def __call__(self, tensor):
            assert str(tensor.device) == device
            outputs = []
            for stride in (8, 16, 32):
                scores = torch.full((1, 1, 2, 2), -20., device=tensor.device)
                if stride == 8:
                    scores[0, 0, 1, 1] = 20.
                outputs.extend((scores, torch.ones((1, 4, 2, 2), device=tensor.device)))
            return outputs
    result = image_detection._predict_yolo(Image.new("RGB", (1280, 1280)), Model(), .5, device=device)
    assert result["boxes"] == [[4., 4., 20., 20.]]
    assert result["scores"] == [1.]
