"""Storage regressions: bounded history, live references and interrupted writes."""

import os
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, TimeoutError
from pathlib import Path
from threading import Event

import pytest
from PIL import Image

from easy_uiauto.mcp import image_storage, interaction_learning, knowledge, scanner, server


def app(tmp_path):
    return knowledge.initialize_app("example", "Example", tmp_path)


def control(directory, color, previous=None, control_id="control-0123456789ab"):
    path, variants, digest = scanner._save_control_images(
        directory, control_id, Image.new("RGB", (8, 8), color), previous or {}
    )
    record = dict(
        id=control_id,
        app_id="example",
        status="observed",
        name="Example",
        image=path.relative_to(directory).as_posix(),
        image_variants=variants,
        image_sha256=digest,
    )
    knowledge.save_control(directory, record)
    return record


def interaction(directory, number, *, color=None, failed=False, legacy=False):
    ident = f"item-{number:03d}"
    if legacy:
        folder = directory / "images" / "interactions" / ident
        folder.mkdir(parents=True)
        image = folder / "before.png"
        Image.new("RGB", (8, 8), number).save(image)
        images = {"before": image.relative_to(directory).as_posix()}
    else:
        image = Image.new("RGB", (8, 8), color if color is not None else number)
        images = {"before": image_storage.save_interaction_image(directory, image)}
    knowledge.save_interaction(
        directory,
        dict(
            id=ident,
            app_id="example",
            command="main.menu.open.click",
            page_id="main",
            status="interference-detected" if failed else "effect-observed",
            images=images,
        ),
    )
    path = directory / "interactions" / f"{ident}.md"
    record, body = knowledge.read_markdown(path)
    # Deliberately stable ordering without relying on filesystem timestamps.
    record["created_at"] = f"2026-10-26T00:00:{number:02d}+00:00"
    knowledge.write_markdown(path, record, body)
    return path, record


def test_unchanged_control_does_not_accumulate_templates(tmp_path):
    directory = app(tmp_path)
    previous = None
    for _ in range(12):
        previous = control(directory, "red", previous)
        image_storage.cleanup_images(directory, dry_run=False)
    assert len(list((directory / "images/controls").rglob("*.png"))) == 2
    assert len(previous["image_variants"]) == 1


def test_control_state_changes_are_bounded_on_disk_and_recent_state_is_refreshed(tmp_path):
    directory = app(tmp_path)
    previous = None
    for color in range(24):
        previous = control(directory, color, previous)
        image_storage.cleanup_images(directory, dry_run=False)
    states = directory / "images/controls" / previous["id"]
    assert len(list(states.glob("*.png"))) == 8
    again = control(directory, 20, previous)
    image_storage.cleanup_images(directory, dry_run=False)
    assert again["image_variants"][-1].endswith(again["image_sha256"] + ".png")
    assert len(list(states.glob("*.png"))) == 8


def test_dry_run_does_not_change_sources_and_apply_protects_cross_record_images(tmp_path):
    directory = app(tmp_path)
    first = control(directory, 0)
    pinned = first["image_variants"][0]
    previous = first
    for color in range(1, 12):
        previous = control(directory, color, previous)
    knowledge.save_page(directory, dict(id="manual", name="Reference", reference_image=pinned))
    source = knowledge.find_control_record(directory, previous["id"])[0]
    before = source.read_bytes()
    preview = image_storage.cleanup_images(directory, keep_control_variants=2)
    assert preview["files"] > 0 and source.read_bytes() == before
    assert (directory / pinned).is_file()
    actual = image_storage.cleanup_images(directory, keep_control_variants=2, dry_run=False)
    assert actual["files"] == preview["files"]
    assert (directory / pinned).is_file()
    assert len(knowledge.read_markdown(source)[0]["image_variants"]) == 2
    assert len(list((directory / "images/controls" / previous["id"]).glob("*.png"))) == 3
    assert (
        image_storage.cleanup_images(directory, keep_control_variants=2, dry_run=False)["files"]
        == 0
    )


@pytest.mark.parametrize(
    "link",
    [
        "![pin](<images/controls/state.png>)",
        '![pin](images/controls/state.png "Evidence")',
        "![[state.png|evidence]]",
        '<img src="images/controls/state.png">',
        '<IMG SRC="images/controls/state.png">',
    ],
)
def test_markdown_html_and_obsidian_image_references_are_preserved(tmp_path, link):
    directory = app(tmp_path)
    image = directory / "images/controls/state.png"
    Image.new("RGB", (8, 8)).save(image)
    # Use a managed hash filename so protection, rather than filename filtering, is tested.
    managed = image.with_name("a" * 64 + ".png")
    image.replace(managed)
    link = link.replace("state.png", managed.name)
    (directory / "notes.md").write_text(link, encoding="utf-8")
    image_storage.cleanup_images(directory, dry_run=False)
    assert managed.is_file()


def test_interaction_snapshots_are_deduplicated_by_content(tmp_path):
    directory = app(tmp_path)
    image = Image.new("RGB", (16, 8), "white")
    snapshot = dict(_target_image=image, _desktop_image=image)
    first = interaction_learning._save_interaction_images(directory, "first", snapshot, snapshot)
    second = interaction_learning._save_interaction_images(directory, "second", snapshot, snapshot)
    assert len(set(first.values()) | set(second.values())) == 1
    assert len(list((directory / "images/interactions").rglob("*.png"))) == 1
    assert (
        image_storage.save_interaction_image(directory, Image.new("RGB", (8, 16), "white"))
        not in first.values()
    )


def test_default_retention_keeps_50_images_and_all_textual_history(tmp_path):
    directory = app(tmp_path)
    for number in range(55):
        interaction(directory, number)
    result = image_storage.cleanup_images(directory, dry_run=False)
    assert len(result["expired_interactions"]) == 5
    assert len(list((directory / "images/interactions").glob("*.png"))) == 50
    history = knowledge.list_interactions(directory, limit=500)
    assert len(history) == 55
    assert sum(bool(item["images"]) for item in history) == 50
    oldest = knowledge.read_markdown(directory / "interactions/item-000.md")[0]
    assert oldest["images"] == {} and oldest["image_retention"]["state"] == "expired"


def test_referenced_record_and_latest_failure_are_pinned_beyond_recent_limit(tmp_path):
    directory = app(tmp_path)
    entries = [interaction(directory, n, failed=n in {1, 2}) for n in range(55)]
    knowledge.save_control(
        directory,
        dict(
            id="effect",
            app_id="example",
            status="observed",
            function_verification={"effect_record": "interactions/item-000.md"},
        ),
    )
    image = entries[3][1]["images"]["before"]
    (directory / "notes.md").write_text(f"![Evidence]({image})", encoding="utf-8")
    result = image_storage.cleanup_images(directory, dry_run=False)
    assert set(result["pinned_interactions"]) == {"item-000", "item-002", "item-003"}
    assert set(result["expired_interactions"]) == {"item-001", "item-004"}
    assert len(list((directory / "images/interactions").glob("*.png"))) == 53


def test_expiring_old_record_does_not_remove_image_shared_by_recent_record(tmp_path):
    directory = app(tmp_path)
    old, _record = interaction(directory, 0, color="red")
    _path, recent = interaction(directory, 1, color="red")
    result = image_storage.cleanup_images(directory, keep_interaction_images=1, dry_run=False)
    assert result["expired_interactions"] == ["item-000"] and result["files"] == 0
    assert knowledge.read_markdown(old)[0]["images"] == {}
    assert (directory / recent["images"]["before"]).is_file()


@pytest.mark.parametrize("link", ["[[item-000]]", "[[interactions/item-000|evidence]]"])
def test_wiki_reference_to_old_interaction_preserves_its_images(tmp_path, link):
    directory = app(tmp_path)
    _path, record = interaction(directory, 0)
    note = directory / "pages/notes.md"
    note.write_text(link, encoding="utf-8")
    result = image_storage.cleanup_images(directory, keep_interaction_images=0, dry_run=False)
    assert result["pinned_interactions"] == ["item-000"]
    assert (directory / record["images"]["before"]).is_file()


def test_latest_failure_uses_stable_control_identity_after_command_rename(tmp_path):
    directory = app(tmp_path)
    for number in range(3):
        path, record = interaction(directory, number, failed=True)
        record.update(control_id="stable", action="click", command=f"renamed-{number}.click")
        knowledge.write_markdown(path, record, "Failure evidence")
    result = image_storage.cleanup_images(directory, keep_interaction_images=0, dry_run=False)
    assert result["pinned_interactions"] == ["item-002"]
    assert set(result["expired_interactions"]) == {"item-000", "item-001"}


def test_legacy_images_are_pruned_without_removing_text_or_page_images(tmp_path):
    directory = app(tmp_path)
    old, record = interaction(directory, 0, legacy=True)
    interaction(directory, 1, legacy=True)
    page = directory / "images/pages/old.png"
    manual = directory / "images/controls/my-notes.png"
    for path in (page, manual):
        Image.new("RGB", (8, 8)).save(path)
    result = image_storage.cleanup_images(directory, keep_interaction_images=1, dry_run=False)
    assert result["files"] == 1 and not (directory / record["images"]["before"]).exists()
    assert old.is_file() and page.is_file() and manual.is_file()


def test_malformed_markdown_aborts_before_deleting_or_rewriting(tmp_path):
    directory = app(tmp_path)
    orphan = directory / "images/controls" / ("f" * 64 + ".png")
    Image.new("RGB", (8, 8)).save(orphan)
    malformed = directory / "controls/bad.md"
    malformed.write_text("---\nimage: broken", encoding="utf-8")
    with pytest.raises(ValueError):
        image_storage.cleanup_images(directory, dry_run=False)
    assert orphan.is_file()


def test_uppercase_markdown_notes_protect_referenced_images(tmp_path):
    directory = app(tmp_path)
    orphan = directory / "images/controls" / ("f" * 64 + ".png")
    Image.new("RGB", (8, 8)).save(orphan)
    (directory / "NOTES.MD").write_text(
        f"![Pinned]({orphan.relative_to(directory).as_posix()})", encoding="utf-8"
    )
    assert image_storage.cleanup_images(directory, dry_run=False)["files"] == 0
    assert orphan.is_file()


def test_unchecked_linked_knowledge_folder_defers_cleanup(tmp_path, monkeypatch):
    directory = app(tmp_path)
    orphan = directory / "images/controls" / ("f" * 64 + ".png")
    Image.new("RGB", (8, 8)).save(orphan)
    linked = directory / "linked-notes"
    linked.mkdir()
    original = image_storage._linked
    monkeypatch.setattr(image_storage, "_linked", lambda path: path == linked or original(path))
    with pytest.raises(ValueError, match="Linked knowledge folders"):
        image_storage.cleanup_images(directory, dry_run=False)
    assert orphan.is_file()


def test_quarantine_failure_images_survive_later_variant_eviction(tmp_path):
    directory = app(tmp_path)
    previous = control(directory, 0)
    failure_image = previous["image_variants"][0]
    previous.update(
        status="quarantined", quarantine=dict(active=True, failure_images=[failure_image])
    )
    knowledge.save_control(directory, previous)
    for color in range(1, 14):
        previous = control(directory, color, previous)
    image_storage.cleanup_images(directory, dry_run=False)
    assert (directory / failure_image).is_file()
    stored = knowledge.find_control_record(directory, previous["id"], include_quarantine=True)[1]
    assert stored["quarantine"]["active"]
    assert len(stored["image_variants"]) == 8


def test_legacy_quarantine_without_image_manifest_is_conservatively_preserved(tmp_path):
    directory = app(tmp_path)
    record = control(directory, 0)
    record.update(status="quarantined", quarantine={"active": True})
    knowledge.save_control(directory, record)
    orphan = directory / "images/controls" / record["id"] / ("f" * 64 + ".png")
    Image.new("RGB", (8, 8)).save(orphan)
    image_storage.cleanup_images(directory, dry_run=False)
    assert orphan.is_file()


def test_linked_image_is_never_deleted_or_followed_outside_application(tmp_path):
    directory = app(tmp_path)
    external = tmp_path / "outside.png"
    Image.new("RGB", (8, 8)).save(external)
    link = directory / "images/controls" / ("e" * 64 + ".png")
    try:
        link.symlink_to(external)
    except OSError:
        pytest.skip("Creating symlinks is unavailable on this Windows account")
    image_storage.cleanup_images(directory, dry_run=False)
    assert link.is_symlink() and external.is_file()


def test_collector_waits_for_pending_image_and_markdown_commit(tmp_path):
    directory = app(tmp_path)
    written, release, collecting = Event(), Event(), Event()

    @image_storage.serialized_learning
    def writer(directory):
        image = image_storage.save_interaction_image(directory, Image.new("RGB", (8, 8)))
        written.set()
        assert release.wait(3)
        knowledge.save_control(
            directory, dict(id="pending", app_id="example", status="observed", image=image)
        )
        return image

    def collect():
        collecting.set()
        return image_storage.cleanup_images(directory, dry_run=False)

    with ThreadPoolExecutor(2) as pool:
        write = pool.submit(writer, directory)
        assert written.wait(2)
        cleanup = pool.submit(collect)
        assert collecting.wait(2)
        try:
            with pytest.raises(TimeoutError):
                cleanup.result(timeout=0.05)
        finally:
            release.set()
        image = write.result(timeout=3)
        assert cleanup.result(timeout=3)["files"] == 0
    assert (directory / image).is_file()


def test_mcp_cleanup_defaults_to_preview(monkeypatch):
    received = {}

    def cleanup(directory, **options):
        received.update(options)
        return dict(ok=True, **options)

    monkeypatch.setattr(server.image_storage, "cleanup_images", cleanup)
    assert '"dry_run": true' in server.cleanup_ui_knowledge_images("example")
    assert received["dry_run"] is True


def test_collector_waits_for_other_process_to_commit_pending_images(tmp_path):
    directory = app(tmp_path)
    ready, release = tmp_path / "ready", tmp_path / "release"
    script = """
import sys, time
from pathlib import Path
from PIL import Image
from easy_uiauto.mcp import image_storage, knowledge
directory, ready, release = map(Path, sys.argv[1:])
with image_storage._vault_lock(directory.parent.parent):
    image = image_storage.save_interaction_image(directory, Image.new('RGB', (8, 8)))
    ready.write_text(image, encoding='utf-8')
    deadline = time.monotonic() + 15
    while not release.exists():
        if time.monotonic() > deadline:
            raise TimeoutError('Waiting for test release')
        time.sleep(0.01)
    knowledge.save_control(directory, dict(id='pending', app_id='example', image=image))
"""
    environment = os.environ.copy()
    environment["PYTHONPATH"] = str(Path(__file__).resolve().parents[1] / "src")
    child = subprocess.Popen(
        [sys.executable, "-c", script, str(directory), str(ready), str(release)],
        env=environment,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    collecting = Event()

    def collect():
        collecting.set()
        return image_storage.cleanup_images(directory, dry_run=False)

    try:
        deadline = time.monotonic() + 10
        while not ready.exists() and child.poll() is None and time.monotonic() < deadline:
            time.sleep(0.01)
        assert ready.is_file()
        with ThreadPoolExecutor(1) as pool:
            cleanup = pool.submit(collect)
            assert collecting.wait(2)
            try:
                with pytest.raises(TimeoutError):
                    cleanup.result(timeout=0.1)
            finally:
                release.touch()
            assert cleanup.result(timeout=10)["files"] == 0
        _out, errors = child.communicate(timeout=5)
        assert child.returncode == 0, errors.decode(errors="replace")
        assert (directory / ready.read_text(encoding="utf-8")).is_file()
    finally:
        release.touch()
        if child.poll() is None:
            child.kill()
        child.communicate(timeout=5)
