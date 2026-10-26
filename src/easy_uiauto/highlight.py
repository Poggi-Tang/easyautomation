"""Keep Tk's complete lifecycle in a private interpreter process."""

from __future__ import annotations

import json
import math
import queue
import subprocess
import sys
import threading
import time
from collections import deque
from pathlib import Path


class HighlightProcess:
    def __init__(self, rect=None, *, color="#FF0000", line_width=2, alpha=1.0,
                 name="easy-uiauto-highlight"):
        self._initial = {"cmd": "init", "rect": rect,
                         "style": {"color": color, "line_width": line_width, "alpha": alpha}}
        self._name = name
        self._process = None
        self._events = queue.Queue()
        self._stderr = deque(maxlen=20)
        self._readers = []
        self._write_lock = threading.Lock()
        self._error = None
        self._forced_termination = False

    @property
    def process(self):
        return self._process

    def _read_stdout(self):
        try:
            for line in self._process.stdout:
                try:
                    value = json.loads(line)
                except ValueError:
                    continue
                if isinstance(value, dict) and value.get("event"):
                    self._events.put(value)
        finally:
            self._events.put({"event": "eof"})

    def _read_stderr(self):
        for line in self._process.stderr:
            self._stderr.append(line.rstrip())

    def _send_raw(self, command):
        with self._write_lock:
            if self._process is None or self._process.poll() is not None:
                raise RuntimeError(self._error or "highlight worker is not running")
            self._process.stdin.write(json.dumps(command, ensure_ascii=False) + "\n")
            self._process.stdin.flush()

    def start(self, timeout=10):
        if not math.isfinite(timeout) or timeout <= 0:
            raise ValueError("highlight timeout must be finite and positive")
        if self._process is not None:
            raise RuntimeError("highlight worker has already been started")
        package_dir = Path(__file__).resolve().parent
        if not (package_dir / "_highlight_worker.py").is_file():
            raise RuntimeError(
                f"Highlight worker is missing from {package_dir}; reinstall easy-uiauto"
            )
        # -E -P ignores Python environment overrides and the unsafe cwd entry,
        # while keeping the interpreter's configured site-packages available.
        # This matters when dependencies are installed in the user site (a
        # common PyCharm setup). Bootstrap the exact package that loaded us,
        # including editable installations, instead of another installed copy.
        bootstrap = ("import runpy, sys; sys.path.insert(0, sys.argv[1]); "
                     "runpy.run_module('easy_uiauto._highlight_worker', run_name='__main__')")
        deadline = time.monotonic() + timeout
        try:
            self._process = subprocess.Popen(
                [sys.executable, "-E", "-P", "-c", bootstrap, str(package_dir.parent)],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                text=True, encoding="utf-8", bufsize=1,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            for stream, target in (("stdout", self._read_stdout), ("stderr", self._read_stderr)):
                reader = threading.Thread(target=target, name=f"{self._name}-{stream}", daemon=True)
                reader.start()
                self._readers.append(reader)
            self._send_raw(self._initial)
            while True:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise TimeoutError("highlight worker start timed out")
                try:
                    event = self._events.get(timeout=remaining)
                except queue.Empty as error:
                    raise TimeoutError("highlight worker start timed out") from error
                if event["event"] == "ready":
                    return
                if event["event"] == "error":
                    raise RuntimeError(event.get("message") or "highlight worker failed")
                if event["event"] == "eof":
                    raise RuntimeError("highlight worker exited before ready")
        except Exception as error:
            status = self.stop(timeout=1)
            self._error = str(error)
            detail = status["stderr"]
            if detail:
                raise type(error)(f"{error}: {detail}") from error
            raise

    def _drain_events(self):
        while True:
            try:
                event = self._events.get_nowait()
            except queue.Empty:
                return
            if event["event"] == "error":
                self._error = event.get("message") or "highlight worker failed"

    def send(self, command):
        self._drain_events()
        if self._error:
            raise RuntimeError(self._error)
        self._send_raw(command)

    def status(self):
        self._drain_events()
        exitcode = self._process.poll() if self._process is not None else None
        stderr = "\n".join(self._stderr) or None
        if not self._error and exitcode not in (None, 0):
            self._error = stderr or f"highlight worker exited with code {exitcode}"
        return {"running": self._process is not None and exitcode is None,
                "stopped": self._process is not None and exitcode is not None,
                "exitcode": exitcode, "error": self._error, "stderr": stderr,
                "forced_termination": self._forced_termination}

    def stop(self, timeout=2):
        process = self._process
        if process is not None:
            if process.poll() is None:
                try:
                    self._send_raw({"cmd": "stop"})
                except (OSError, RuntimeError):
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
            for reader in self._readers:
                reader.join(timeout=1)
            for stream in (process.stdin, process.stdout, process.stderr):
                if stream is not None:
                    try:
                        stream.close()
                    except OSError:
                        pass  # Closing stdin can flush into an already exited child.
        return self.status()
