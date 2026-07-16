from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
from typing import Any


class HighlightProcess:
    """Run the Tk overlay in a dedicated interpreter and exchange JSON commands."""

    def __init__(
        self,
        rect: dict[str, int] | None = None,
        *,
        color: str = "#FF0000",
        line_width: int = 2,
        alpha: float = 1.0,
        name: str = "easy-uiauto-highlight",
    ) -> None:
        self._initial = {
            "cmd": "init",
            "rect": dict(rect) if rect else None,
            "style": {
                "color": color,
                "line_width": max(1, int(line_width)),
                "alpha": max(0.05, min(float(alpha), 1.0)),
            },
        }
        self._name = name
        self._process: subprocess.Popen[str] | None = None
        self._events: queue.Queue[dict[str, Any]] = queue.Queue()
        self._stdout_noise: list[str] = []
        self._stderr: list[str] = []
        self._reader_threads: list[threading.Thread] = []
        self._write_lock = threading.Lock()
        self._error: str | None = None
        self._started = False
        self._forced_termination = False

    @property
    def process(self) -> subprocess.Popen[str] | None:
        return self._process

    def _read_stdout(self) -> None:
        process = self._process
        if process is None or process.stdout is None:
            return
        for line in process.stdout:
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                self._stdout_noise.append(line.rstrip())
                continue
            if isinstance(event, dict) and event.get("event"):
                self._events.put(event)
            else:
                self._stdout_noise.append(line.rstrip())

    def _read_stderr(self) -> None:
        process = self._process
        if process is None or process.stderr is None:
            return
        for line in process.stderr:
            self._stderr.append(line.rstrip())

    def _send_raw(self, command: dict[str, Any]) -> None:
        process = self._process
        if process is None or process.stdin is None or process.poll() is not None:
            raise RuntimeError(self._error or "highlight worker is not running")
        payload = json.dumps(command, ensure_ascii=False)
        with self._write_lock:
            process.stdin.write(payload + "\n")
            process.stdin.flush()

    def _drain_events(self) -> None:
        while True:
            try:
                event = self._events.get_nowait()
            except queue.Empty:
                break
            if event.get("event") == "error":
                self._error = str(event.get("message") or "highlight worker error")

    def start(self, timeout: float = 10) -> None:
        if self._started:
            raise RuntimeError("highlight worker has already been started")
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        self._process = subprocess.Popen(
            [sys.executable, "-I", "-m", "easy_uiauto._highlight_worker"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            bufsize=1,
            env=os.environ.copy(),
            creationflags=creationflags,
        )
        self._started = True
        self._reader_threads = [
            threading.Thread(
                target=self._read_stdout,
                name=f"{self._name}-stdout",
                daemon=True,
            ),
            threading.Thread(
                target=self._read_stderr,
                name=f"{self._name}-stderr",
                daemon=True,
            ),
        ]
        for thread in self._reader_threads:
            thread.start()
        try:
            self._send_raw(self._initial)
            event = self._events.get(timeout=timeout)
        except queue.Empty as exc:
            status = self.stop()
            detail = status.get("stderr") or "worker produced no ready event"
            raise TimeoutError(f"highlight worker start timed out: {detail}") from exc
        except Exception:
            self.stop()
            raise
        if event.get("event") != "ready":
            self._error = str(event.get("message") or "highlight worker failed to start")
            self.stop()
            raise RuntimeError(self._error)

    def send(self, command: dict[str, Any]) -> None:
        self._drain_events()
        if self._error is not None:
            raise RuntimeError(self._error)
        self._send_raw(command)

    def status(self) -> dict[str, Any]:
        self._drain_events()
        process = self._process
        running = bool(process is not None and process.poll() is None)
        exitcode = process.poll() if process is not None else None
        stderr = "\n".join(self._stderr[-10:]) or None
        stdout_noise = "\n".join(self._stdout_noise[-10:]) or None
        if self._error is None and exitcode not in (None, 0):
            self._error = stderr or stdout_noise or f"highlight worker exited with code {exitcode}"
        return {
            "running": running,
            "stopped": bool(self._started and not running),
            "forced_termination": self._forced_termination,
            "exitcode": exitcode,
            "error": self._error,
            "stderr": stderr,
            "stdout_noise": stdout_noise,
        }

    def stop(self, timeout: float = 5) -> dict[str, Any]:
        process = self._process
        if process is not None and process.poll() is None:
            try:
                self._send_raw({"cmd": "stop"})
            except (BrokenPipeError, OSError, RuntimeError):
                pass
            try:
                process.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                self._forced_termination = True
                process.terminate()
                try:
                    process.wait(timeout=2)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=2)
        if process is not None:
            for stream in (process.stdin, process.stdout, process.stderr):
                if stream is not None:
                    try:
                        stream.close()
                    except OSError:
                        pass
        for thread in self._reader_threads:
            thread.join(timeout=1)
        return self.status()
