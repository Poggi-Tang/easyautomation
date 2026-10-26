"""Bounded image retention, preserving Markdown knowledge and referenced evidence."""

from __future__ import annotations

import functools
import hashlib
import inspect
import io
import os
import re
import threading
import time
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import unquote
from uuid import uuid4

CONTROL_VARIANTS_ENV = "EASY_UIAUTO_CONTROL_IMAGE_VARIANTS"
INTERACTION_IMAGES_ENV = "EASY_UIAUTO_INTERACTION_IMAGE_RECORDS"
_LOCKS: dict[str, threading.RLock] = {}
_LOCKS_GUARD = threading.Lock()
_LOCK_DEPTH = threading.local()
_HASH_IMAGE = re.compile(r"[0-9a-f]{64}\.png", re.I)
_CONTROL_IMAGE = re.compile(r"(?:control|visual)-[0-9a-f]{12}\.png", re.I)
_INTERACTION_ID = re.compile(r"interaction-[0-9a-f]{12}", re.I)
_LEGACY_IMAGES = {"before.png", "after.png", "desktop-before.png", "desktop-after.png"}


def _limit(value, env: str, default: int, minimum: int) -> int:
    if value is None:
        value = os.environ.get(env, str(default))
    number = int(value)
    if not minimum <= number <= 10000:
        raise ValueError(f"{env} must be between {minimum} and 10000")
    return number


def control_variant_limit() -> int:
    return _limit(None, CONTROL_VARIANTS_ENV, 8, 1)


@contextmanager
def _vault_lock(root: Path):
    """Serialize cooperating writers across threads and MCP processes."""
    key = str(root.resolve()).casefold()
    with _LOCKS_GUARD:
        lock = _LOCKS.setdefault(key, threading.RLock())
    with lock:
        held = getattr(_LOCK_DEPTH, "held", None)
        if held is None:
            held = _LOCK_DEPTH.held = set()
        if key in held:
            yield
            return
        folder = root.resolve() / ".easy_uiauto"
        path = folder / "images.lock"
        if _linked(folder) or _linked(path):
            raise ValueError("The image storage lock must not be a symlink or junction")
        folder.mkdir(parents=True, exist_ok=True)
        with path.open("a+b") as stream:
            if stream.seek(0, os.SEEK_END) == 0:
                stream.write(b"\0")
                stream.flush()
            stream.seek(0)
            if os.name == "nt":
                import msvcrt

                while True:
                    try:
                        msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
                        break
                    except OSError as error:
                        if error.errno not in {11, 13, 36}:
                            raise
                        time.sleep(0.05)
            else:
                import fcntl

                fcntl.flock(stream, fcntl.LOCK_EX)
            held.add(key)
            try:
                yield
            finally:
                held.remove(key)
                stream.seek(0)
                if os.name == "nt":
                    msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    fcntl.flock(stream, fcntl.LOCK_UN)


def serialized_learning(function):
    """Keep image writes + record commits separate from same-vault collection."""
    signature = inspect.signature(function)

    @functools.wraps(function)
    def wrapped(*args, **kwargs):
        from . import knowledge

        arguments = signature.bind(*args, **kwargs).arguments
        directory = arguments.get("directory")
        root = (
            Path(directory).parent.parent
            if directory is not None
            else (arguments.get("root") or knowledge.vault_root())
        )
        with _vault_lock(Path(root)):
            return function(*args, **kwargs)

    return wrapped


def _linked(path: Path) -> bool:
    return path.is_symlink() or getattr(path, "is_junction", lambda: False)()


def _safe_path(directory: Path, path: Path) -> Path | None:
    """Never follow a symlink/junction or delete a resolved path outside this app."""
    root = directory.resolve()
    if not path.is_absolute():
        path = directory / path
    lexical = Path(os.path.abspath(path))
    if not lexical.is_relative_to(root):
        return None
    current = lexical
    while current != root:
        if _linked(current):
            return None
        current = current.parent
    resolved = lexical.resolve()
    return resolved if resolved.is_relative_to(root) else None


def _reference(directory: Path, value: str, origin: Path, cache=None) -> Path | None:
    value = unquote(value.strip().strip("<>"))
    if value.startswith(("http:", "https:", "data:", "file:")):
        return None
    value = value.split("#", 1)[0].split("?", 1)[0]
    if Path(value).suffix.casefold() not in {".png", ".jpg", ".jpeg", ".webp", ".md"}:
        return None
    key = (directory, origin, value)
    if cache is not None and key in cache:
        return cache[key]
    try:
        path = Path(value.replace("\\", "/"))
        candidate = (path if path.is_absolute() else origin / path).resolve()
        result = candidate if candidate.is_relative_to(directory.resolve()) else None
    except (OSError, ValueError):
        result = None
    if cache is not None:
        cache[key] = result
    return result


def _references(directory: Path, value, origin: Path, cache=None) -> set[Path]:
    result = set()
    if isinstance(value, dict):
        for item in value.values():
            result.update(_references(directory, item, origin, cache))
    elif isinstance(value, (list, tuple)):
        for item in value:
            result.update(_references(directory, item, origin, cache))
    elif isinstance(value, str):
        path = _reference(directory, value, origin, cache)
        if path:
            result.add(path)
        # Markdown/HTML/code links, including percent-encoded and angle-bracket paths.
        for match in re.finditer(
            r"\]\((<[^>]+>|[^)]+)\)|(?:src|href)=[\"']([^\"']+)|`([^`]+)`", value, re.I
        ):
            raw = next(part for part in match.groups() if part).strip()
            if raw.startswith("<") and ">" in raw:
                raw = raw[1 : raw.index(">")]
            else:
                raw = re.split(r"\s+[\"']", raw, maxsplit=1)[0]
            candidate = _reference(directory, raw, origin, cache)
            if candidate:
                result.add(candidate)
        for match in re.finditer(r"\[\[([^\]|]+)(?:\|[^\]]*)?\]\]", value):
            raw = match.group(1).split("#", 1)[0].strip()
            if not Path(raw).suffix:
                raw += ".md"
            for base in (origin, directory):
                candidate = _reference(directory, raw, base, cache)
                if candidate:
                    result.add(candidate)
            if "/" not in raw and "\\" not in raw:
                pattern = f"**/{raw}" if raw.endswith(".md") else f"images/**/{raw}"
                for image in directory.glob(pattern):
                    safe = _reference(directory, str(image), directory, cache)
                    if safe and safe.is_file():
                        result.add(safe)
    return result


def _records(directory: Path):
    from . import knowledge

    records = []
    for parent, folders, files in os.walk(directory, followlinks=False):
        # An unchecked linked knowledge folder could contain a live image reference.
        folders[:] = [name for name in folders if name not in {"images", ".easy_uiauto"}]
        if any(_linked(Path(parent) / name) for name in folders):
            raise ValueError("Linked knowledge folders cannot be checked safely")
        for name in sorted(files):
            if Path(name).suffix.casefold() != ".md":
                continue
            path = _safe_path(directory, Path(parent) / name)
            if path is None:
                raise ValueError(f"Linked Markdown cannot be checked safely: {name}")
            text = path.read_text(encoding="utf-8")
            if text.lstrip("\ufeff").startswith("---\n"):
                metadata, body = knowledge.read_markdown(path)
            else:
                metadata, body = {}, text
            if metadata.get("kind") == "control" and (
                not isinstance(metadata.get("image_variants", []), list)
                or not all(isinstance(value, str) for value in metadata.get("image_variants", []))
            ):
                raise ValueError(f"Malformed image_variants in {name}")
            if metadata.get("kind") == "interaction" and not isinstance(
                metadata.get("images", {}), dict
            ):
                raise ValueError(f"Malformed interaction images in {name}")
            records.append([path, metadata, body])
    return records


def _all_refs(directory: Path, records, cache=None) -> set[Path]:
    refs = set()
    for path, record, body in records:
        refs.update(_references(directory, record, directory, cache))
        refs.update(_references(directory, body, path.parent, cache))
    return refs


def save_interaction_image(directory: Path, image) -> str:
    """Identical PNG content shares one immutable image across all interactions."""
    if _linked(directory):
        raise ValueError("The application directory must not be a symlink or junction")
    directory = directory.resolve()
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    content = buffer.getvalue()
    digest = hashlib.sha256(content).hexdigest()
    path = directory / "images" / "interactions" / f"{digest}.png"
    if _safe_path(directory, path) is None:
        raise ValueError("Interaction image path leaves the application directory")
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != digest:
        temporary = path.with_name(f".{digest}-{uuid4().hex}.tmp")
        try:
            temporary.write_bytes(content)
            temporary.replace(path)
        finally:
            temporary.unlink(missing_ok=True)
    return path.relative_to(directory).as_posix()


def _managed_images(directory: Path, records):
    controls = directory / "images" / "controls"
    interactions = directory / "images" / "interactions"
    interaction_ids = {
        str(record.get("id")) for _p, record, _b in records if record.get("kind") == "interaction"
    }
    for folder in (controls, interactions):
        if _safe_path(directory, folder) is None or not folder.is_dir():
            continue
        for child in folder.iterdir():
            if _safe_path(directory, child) is None:
                continue
            if child.is_file() and (
                _HASH_IMAGE.fullmatch(child.name)
                or folder == controls
                and _CONTROL_IMAGE.fullmatch(child.name)
            ):
                yield child
            elif child.is_dir() and (
                folder == controls
                or child.name in interaction_ids
                or _INTERACTION_ID.fullmatch(child.name)
            ):
                for image in child.iterdir():
                    if (
                        _safe_path(directory, image) is not None
                        and image.is_file()
                        and (
                            _HASH_IMAGE.fullmatch(image.name)
                            or (folder == interactions and image.name in _LEGACY_IMAGES)
                        )
                    ):
                        yield image


@serialized_learning
def cleanup_images(
    directory: Path,
    *,
    keep_control_variants: int | None = None,
    keep_interaction_images: int | None = None,
    dry_run: bool = True,
) -> dict:
    """Collect only managed unreferenced PNGs; retain textual interaction history."""
    from . import knowledge

    if _linked(directory):
        raise ValueError("The application directory must not be a symlink or junction")
    directory = directory.resolve()
    if not (directory / "app.md").is_file():
        raise ValueError("An initialized application knowledge directory is required")
    variants_limit = _limit(keep_control_variants, CONTROL_VARIANTS_ENV, 8, 1)
    history_limit = _limit(keep_interaction_images, INTERACTION_IMAGES_ENV, 50, 0)
    records = _records(directory)  # Any unreadable source aborts before mutation.
    reference_cache = {}
    before_refs = _all_refs(directory, records, reference_cache)
    interactions = [item for item in records if item[1].get("kind") == "interaction"]
    interactions.sort(
        key=lambda item: (item[1].get("created_at") or item[1].get("updated_at", ""), item[0].name),
        reverse=True,
    )
    evidence_refs = set()
    for path, record, body in records:
        evidence = {
            key: value
            for key, value in record.items()
            if record.get("kind") != "interaction" or key != "images"
        }
        evidence_refs.update(_references(directory, evidence, directory, reference_cache))
        evidence_refs.update(_references(directory, body, path.parent, reference_cache))
    latest_failures = set()
    failed_commands = set()
    for path, record, _body in interactions:
        failed = (
            record.get("status") in {"failed", "interference-detected"}
            or record.get("recovery", {}).get("ok") is False
            or record.get("execution", {}).get("ok") is False
        )
        anchor = record.get("control_anchor") or record.get("control_id") or record.get("command")
        command = (
            record.get("page_id"),
            repr(anchor),
            record.get("action") or str(record.get("command", "")).rsplit(".", 1)[-1],
        )
        if failed and record.get("images") and command not in failed_commands:
            latest_failures.add(path)
            failed_commands.add(command)
    with_images = [item for item in interactions if item[1].get("images")]
    recent = {path for path, _record, _body in with_images[:history_limit]}
    changed = []
    expired = []
    pinned = []
    legacy_quarantine_dirs = set()
    for path, record, body in records:
        if record.get("kind") == "control":
            variants = list(dict.fromkeys(record.get("image_variants", [])))[-variants_limit:]
            if variants != record.get("image_variants", []):
                record["image_variants"] = variants
                changed.append((path, record, body))
            if knowledge.runtime_quarantined(record) and not record.get("quarantine", {}).get(
                "failure_images"
            ):
                folder = _safe_path(
                    directory, directory / "images" / "controls" / str(record.get("id"))
                )
                if folder:
                    legacy_quarantine_dirs.add(folder)
        if record.get("kind") != "interaction" or not record.get("images"):
            continue
        image_refs = _references(directory, record["images"], directory, reference_cache)
        is_pinned = (
            path in evidence_refs
            or bool(image_refs & evidence_refs)
            or path in latest_failures
            or record.get("retain_images") is True
        )
        if is_pinned:
            pinned.append(record.get("id"))
        if path in recent or is_pinned:
            continue
        record["image_retention"] = dict(
            state="expired",
            expired_at=knowledge.utc_now(),
            reason="outside retained recent or referenced evidence",
            roles=sorted(record["images"]),
        )
        record["images"] = {}
        changed.append((path, record, body))
        expired.append(record.get("id"))
    refs = _all_refs(directory, records, reference_cache)
    candidates = [
        path
        for path in _managed_images(directory, records)
        if path not in refs and path.parent not in legacy_quarantine_dirs
    ]
    reclaimed = sum(path.stat().st_size for path in candidates)
    report = dict(
        ok=True,
        dry_run=dry_run,
        keep_control_variants=variants_limit,
        keep_interaction_images=history_limit,
        pinned_interactions=pinned,
        expired_interactions=expired,
        updated_records=len(changed),
        files=len(candidates),
        bytes=reclaimed,
        paths=[path.relative_to(directory).as_posix() for path in candidates],
    )
    if dry_run:
        return report
    if not changed and not candidates:
        report.update(errors=[], references_checked=len(before_refs))
        return report
    # Commit references first; a crash can leave extra PNGs, never broken live refs.
    for path, record, body in changed:
        knowledge.write_markdown(path, record, body)
    # Re-read to honor source changes before deleting. Writers share the vault lock.
    live_refs = _all_refs(directory, _records(directory))
    removed = []
    errors = []
    actual_bytes = 0
    for path in candidates:
        if path in live_refs or _safe_path(directory, path) is None:
            continue
        try:
            size = path.stat().st_size
            path.unlink()
            removed.append(path.relative_to(directory).as_posix())
            actual_bytes += size
            parent = path.parent
            if parent.name not in {"controls", "interactions"} and not any(parent.iterdir()):
                parent.rmdir()
        except OSError as error:
            errors.append(str(error))
    knowledge.rebuild_index(directory)
    report.update(
        files=len(removed),
        bytes=actual_bytes,
        paths=removed,
        errors=errors,
        references_checked=len(before_refs),
    )
    return report
