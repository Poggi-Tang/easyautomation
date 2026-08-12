"""Tests for bundled Codex skill installation."""

from __future__ import annotations

from easy_uiauto.mcp import skill_installation


def test_install_and_update_bundled_codex_skills(monkeypatch, tmp_path) -> None:
    source = tmp_path / "source"
    destination = tmp_path / "codex" / "skills"
    for name in skill_installation.SKILL_NAMES:
        skill = source / name
        (skill / "agents").mkdir(parents=True)
        (skill / "SKILL.md").write_text(f"---\nname: {name}\n---\n", encoding="utf-8")
        (skill / "agents" / "openai.yaml").write_text("interface: {}\n", encoding="utf-8")
    stale = destination / skill_installation.SKILL_NAMES[0]
    stale.mkdir(parents=True)
    (stale / "stale.txt").write_text("old", encoding="utf-8")
    monkeypatch.setattr(skill_installation, "bundled_skills_dir", lambda: source)
    monkeypatch.setattr(skill_installation, "codex_skills_dir", lambda: destination)

    output = skill_installation.install_codex_skills()

    assert "Installed Codex skills" in output
    assert not (stale / "stale.txt").exists()
    for name in skill_installation.SKILL_NAMES:
        assert (destination / name / "SKILL.md").is_file()


def test_uninstall_removes_only_easy_uiauto_skills(monkeypatch, tmp_path) -> None:
    destination = tmp_path / "skills"
    for name in (*skill_installation.SKILL_NAMES, "unrelated"):
        (destination / name).mkdir(parents=True)
    monkeypatch.setattr(skill_installation, "codex_skills_dir", lambda: destination)

    skill_installation.uninstall_codex_skills()

    assert (destination / "unrelated").is_dir()
    for name in skill_installation.SKILL_NAMES:
        assert not (destination / name).exists()
