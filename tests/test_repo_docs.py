from pathlib import Path


def test_readme_and_agents_document_verification_gate() -> None:
    readme = Path("README.md").read_text()
    agents = Path("AGENTS.md").read_text()

    for command in [
        "rtk uv run pytest",
        "rtk uv run ruff check .",
        "rtk uv run ruff format --check .",
    ]:
        assert command in readme
        assert command in agents


def test_agents_documents_skill_symlink_rule() -> None:
    agents = Path("AGENTS.md").read_text()

    assert "skills/claude/SKILL.md" in agents
    assert "symlink" in agents
    assert "skills/codex/SKILL.md" in agents


def test_readme_documents_next_milestone_commands() -> None:
    readme = Path("README.md").read_text()
    for command in [
        "sheetops cells read",
        "sheetops styles assert",
        "sheetops formulas scan",
        "sheetops normalize preview",
        "sheetops recalc",
        "sheetops diff",
        "sheetops reconcile",
    ]:
        assert command in readme


def test_readme_documents_libreoffice_doctor_and_setup_script() -> None:
    readme = Path("README.md").read_text()

    assert "sheetops doctor libreoffice" in readme
    assert "scripts/setup-libreoffice" in readme
