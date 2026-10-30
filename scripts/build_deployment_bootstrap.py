"""Produce the standalone setup launchers attached to a GitHub release."""

from pathlib import Path
import argparse
import hashlib
import shutil
import re
import tomllib


def main():
    root = Path(__file__).resolve().parents[1]
    source = (root / "src/easy_uiauto/deployment.py").read_text(encoding="utf-8")
    digest = hashlib.sha256(source.encode("utf-8")).hexdigest()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=root / "dist/deployment")
    destination = parser.parse_args().output_dir
    destination.mkdir(parents=True, exist_ok=True)
    text = (root / "scripts/full_setup.ps1").read_text(encoding="utf-8")
    version = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    text = re.sub(r'\$packageVersion = "[^"]+"', f'$packageVersion = "{version}"', text, count=1)
    # Replace the value only, leaving the unbuilt-template guard unchanged.
    text = text.replace('$deploymentSha256 = "__DEPLOYMENT_SHA256__"', f'$deploymentSha256 = "{digest}"')
    (destination / "easy-uiauto-setup.ps1").write_text(text, encoding="utf-8-sig")
    for name in ("install-full.cmd", "download-offline.cmd"):
        shutil.copy2(root / name, destination / name)
    print("Standalone deployment launchers generated")


if __name__ == "__main__":
    main()
