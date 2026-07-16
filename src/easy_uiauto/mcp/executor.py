from __future__ import annotations

import queue
import threading
from collections.abc import Callable
from concurrent.futures import Future
from concurrent.futures import TimeoutError as FutureTimeoutError
from dataclasses import dataclass
from typing import Any

from easy_uiauto.utils import ensure_uiautomation_thread


@dataclass(slots=True)
class _Request:
    method: str
    args: tuple[Any, ...]
    kwargs: dict[str, Any]
    future: Future


class UIAutomationExecutor:
    """Serialize every UIAutomation/COM operation onto one dedicated thread."""

    def __init__(self, backend_factory: Callable[[], Any], *, name: str = "easy-uiauto-mcp"):
        self._backend_factory = backend_factory
        self._queue: queue.Queue[_Request | None] = queue.Queue()
        self._ready = threading.Event()
        self._startup_error: BaseException | None = None
        self._backend: Any = None
        self._thread = threading.Thread(target=self._run, name=name, daemon=True)
        self._thread.start()
        if not self._ready.wait(timeout=10):
            raise TimeoutError("UIAutomation worker thread did not start")
        if self._startup_error is not None:
            raise RuntimeError("UIAutomation worker thread failed to start") from self._startup_error

    @property
    def thread_id(self) -> int | None:
        return self._thread.ident

    def _run(self) -> None:
        try:
            ensure_uiautomation_thread()
            self._backend = self._backend_factory()
        except BaseException as exc:
            self._startup_error = exc
            self._ready.set()
            return
        self._ready.set()
        while True:
            request = self._queue.get()
            if request is None:
                break
            if not request.future.set_running_or_notify_cancel():
                continue
            try:
                method = getattr(self._backend, request.method)
                request.future.set_result(method(*request.args, **request.kwargs))
            except BaseException as exc:
                request.future.set_exception(exc)
        close = getattr(self._backend, "close", None)
        if callable(close):
            close()

    def call(self, method: str, *args: Any, timeout: float = 30, **kwargs: Any) -> Any:
        if not self._thread.is_alive():
            raise RuntimeError("UIAutomation worker thread is not running")
        future: Future = Future()
        self._queue.put(_Request(method=method, args=args, kwargs=kwargs, future=future))
        try:
            return future.result(timeout=timeout)
        except FutureTimeoutError:
            if future.cancel():
                raise
            # The operation has already started and COM/UI input cannot be safely cancelled.
            # Wait for the authoritative result instead of returning while a mutation continues.
            return future.result()

    def close(self) -> None:
        if not self._thread.is_alive():
            return
        self._queue.put(None)
        if self._thread is not threading.current_thread():
            self._thread.join(timeout=5)

    def __enter__(self) -> UIAutomationExecutor:
        return self

    def __exit__(self, *_exc_info: object) -> None:
        self.close()
