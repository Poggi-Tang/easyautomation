from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from typing import Any

from easy_uiauto.highlight import HighlightProcess
from easy_uiauto.record import RecordThread


@dataclass(slots=True)
class RecordingSession:
    recorder: RecordThread
    session_id: str = field(default_factory=lambda: uuid.uuid4().hex)

    @classmethod
    def start(cls) -> RecordingSession:
        session = cls(recorder=RecordThread())
        session.recorder.start()
        session.recorder.started.wait(timeout=10)
        if session.recorder.run_error is not None:
            session._stop_and_join()
            raise RuntimeError("recording listener failed to start") from session.recorder.run_error
        if not session.recorder.started.is_set():
            session._stop_and_join()
            raise TimeoutError("recording listener start timed out")
        return session

    def _stop_and_join(self) -> None:
        self.recorder.stop()
        self.recorder.join(timeout=5)
        if self.recorder.is_alive():
            self.recorder.cleanup_errors.append("录制协调线程未在超时内退出")

    def status(self, *, include_actions: bool = True) -> dict[str, Any]:
        with self.recorder._state_lock:
            actions = [dict(action) for action in self.recorder.actions_data]
            return {
                "session_id": self.session_id,
                "running": bool(
                    self.recorder.is_alive()
                    and self.recorder.running
                    and not self.recorder._stopped
                ),
                "stopped": bool(self.recorder._stopped),
                "action_count": len(actions),
                "actions": actions if include_actions else None,
                "cleanup_errors": list(self.recorder.cleanup_errors),
                "run_error": str(self.recorder.run_error) if self.recorder.run_error else None,
            }

    def stop(self) -> dict[str, Any]:
        self._stop_and_join()
        return self.status()


class HighlightSession:
    def __init__(
        self,
        rect: dict[str, int],
        *,
        color: str = "#FF0000",
        line_width: int = 2,
        alpha: float = 1.0,
    ) -> None:
        self.session_id = uuid.uuid4().hex
        self._worker = HighlightProcess(
            rect,
            color=color,
            line_width=line_width,
            alpha=alpha,
            name=f"easy-uiauto-highlight-{self.session_id[:8]}",
        )

    def start(self) -> None:
        self._worker.start()

    def update_rect(self, rect: dict[str, int]) -> None:
        self._worker.send({"cmd": "rect", "rect": dict(rect)})

    def update_style(self, **style: Any) -> None:
        allowed = {key: value for key, value in style.items() if value is not None}
        self._worker.send({"cmd": "style", "style": allowed})

    def status(self) -> dict[str, Any]:
        return {"session_id": self.session_id, **self._worker.status()}

    def stop(self) -> dict[str, Any]:
        return {"session_id": self.session_id, **self._worker.stop()}
