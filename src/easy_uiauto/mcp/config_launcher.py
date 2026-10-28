"""Launch the bundled Windows configuration app, with a source-checkout fallback."""

from __future__ import annotations

import os
import subprocess
from pathlib import Path


def main():
    if os.name != "nt":
        raise RuntimeError("The configuration app supports Windows only")
    executable = Path(__file__).resolve().parents[1] / "bin" / "easy-uiauto-config.exe"
    if executable.is_file():
        subprocess.Popen([str(executable)], close_fds=True)
        return
    # Source checkouts can still open the editor before the first EXE build.
    from .config_gui import main as run_source

    run_source()


if __name__ == "__main__":
    main()
