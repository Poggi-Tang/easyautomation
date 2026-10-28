"""Check that the wheel and sdist carry the exact compiled Windows GUI EXE."""

from __future__ import annotations

import argparse
import hashlib
import struct
import tarfile
import tomllib
import zipfile
from pathlib import Path


def main():
    root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dist-dir", type=Path, default=root / "dist")
    args = parser.parse_args()
    version = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    relative = "easy_uiauto/bin/easy-uiauto-config.exe"
    executable = (root / "src" / relative).read_bytes()
    if executable[:2] != b"MZ":
        raise RuntimeError("The bundled configuration EXE is missing or invalid")
    offset = struct.unpack_from("<I", executable, 0x3C)[0]
    if executable[offset:offset + 4] != b"PE\x00\x00" or struct.unpack_from(
        "<H", executable, offset + 4
    )[0] != 0x8664:
        raise RuntimeError("The bundled configuration EXE must target Windows x64")
    wheel = args.dist_dir / f"easy_uiauto-{version}-py3-none-win_amd64.whl"
    sdist = args.dist_dir / f"easy_uiauto-{version}.tar.gz"
    with zipfile.ZipFile(wheel) as archive:
        if archive.read(relative) != executable:
            raise RuntimeError("Wheel does not contain the compiled configuration EXE")
        entries = archive.read(f"easy_uiauto-{version}.dist-info/entry_points.txt").decode("utf-8")
        if "easy_uiauto_config = easy_uiauto.mcp.config_launcher:main" not in entries:
            raise RuntimeError("Wheel is missing its configuration launcher")
    with tarfile.open(sdist) as archive:
        member = archive.extractfile(f"easy_uiauto-{version}/src/{relative}")
        if member is None or member.read() != executable:
            raise RuntimeError("Source package does not contain the compiled configuration EXE")
    print("Bundled configuration EXE verified in wheel and sdist")
    print("SHA256:", hashlib.sha256(executable).hexdigest())


if __name__ == "__main__":
    main()
