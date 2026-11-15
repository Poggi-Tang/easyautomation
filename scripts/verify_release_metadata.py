"""Reject a release whose source, lockfile, documentation or tag versions disagree."""

import argparse
import ast
from pathlib import Path
import tomllib


def verify(root: Path, tag: str = "") -> str:
    project = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))["project"]
    version = project["version"]
    tree = ast.parse((root / "src/easy_uiauto/__init__.py").read_text(encoding="utf-8"))
    declared = next(ast.literal_eval(node.value) for node in tree.body
                    if isinstance(node, ast.Assign)
                    and any(isinstance(target, ast.Name) and target.id == "__version__"
                            for target in node.targets))
    assert declared == version, "Package version mismatch"
    lock = tomllib.loads((root / "uv.lock").read_text(encoding="utf-8"))
    locked = next(package for package in lock["package"] if package["name"] == "easy-uiauto")
    assert locked["version"] == version, "Lockfile version mismatch"
    for filename in ("README.md", "README.zh-CN.md"):
        assert f"**{version}**" in (root / filename).read_text(encoding="utf-8"), f"{filename} version mismatch"
    assert f"## [{version}]" in (root / "CHANGELOG.md").read_text(encoding="utf-8"), "Missing release notes"
    assert project["license"] == "LicenseRef-easy-uiauto-Personal-Enterprise", "License metadata mismatch"
    assert not any("MIT License" in item for item in project.get("classifiers", [])), "Unexpected MIT classifier"
    assert (root / "LICENSE").is_file(), "Missing license"
    if tag:
        assert tag == f"v{version}", "Tag/version mismatch"
    return version


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tag", default="")
    args = parser.parse_args()
    print("Release metadata verified: " + verify(Path(__file__).resolve().parents[1], args.tag))
