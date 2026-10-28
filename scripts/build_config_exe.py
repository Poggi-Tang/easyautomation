"""Build the standalone Windows settings app without importing automation backends."""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


def main():
    root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=root / "dist")
    parser.add_argument("--work-dir", type=Path, default=root / "build" / "config-exe")
    parser.add_argument("--bundle-package", action="store_true",
                        help="Also place the compiled EXE inside the Python package")
    args = parser.parse_args()
    if os.name != "nt":
        parser.error("Build this Windows EXE on Windows")
    work = args.work_dir.resolve()
    work.mkdir(parents=True, exist_ok=True)
    source = root / "src" / "easy_uiauto" / "mcp"
    # Running the file directly keeps easy_uiauto.__init__ and its UIA, mouse,
    # screenshot and recording dependencies out of this small configuration EXE.
    subprocess.run([
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
        "--onefile", "--windowed", "--noupx", "--icon", "NONE",
        "--name", "easy-uiauto-config",
        "--distpath", str(args.output_dir.resolve()),
        "--workpath", str(work / "work"), "--specpath", str(work),
        "--paths", str(source), str(source / "config_gui.py"),
    ], check=True)
    executable = args.output_dir.resolve() / "easy-uiauto-config.exe"
    print(executable)
    if args.bundle_package:
        bundled = root / "src" / "easy_uiauto" / "bin" / executable.name
        bundled.parent.mkdir(parents=True, exist_ok=True)
        if executable.resolve() != bundled.resolve():
            shutil.copy2(executable, bundled)
        print(bundled)


if __name__ == "__main__":
    main()
