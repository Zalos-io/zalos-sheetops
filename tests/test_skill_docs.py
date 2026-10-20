import re
from pathlib import Path


def test_claude_skill_is_symlink_to_codex_skill() -> None:
    path = Path("skills/claude/SKILL.md")

    assert path.is_symlink()
    assert path.readlink() == Path("../codex/SKILL.md")


def test_agent_skills_document_token_bounding_knobs() -> None:
    for path in [Path("skills/codex/SKILL.md"), Path("skills/claude/SKILL.md")]:
        text = _skill_bundle(path)
        assert "maxRows" in text
        assert "maxFormulaCells" in text
        assert '"include"' in text


def test_agent_skills_include_next_milestone_commands() -> None:
    for path in [Path("skills/codex/SKILL.md"), Path("skills/claude/SKILL.md")]:
        text = _skill_bundle(path)
        assert "sheetops cells read" in text
        assert "sheetops styles assert" in text
        assert "sheetops formulas scan" in text
        assert "sheetops normalize preview" in text
        assert "sheetops recalc" in text
        assert "sheetops diff" in text
        assert "sheetops reconcile" in text


def test_skills_document_libreoffice_doctor_before_recalc() -> None:
    for path in [Path("skills/codex/SKILL.md"), Path("skills/claude/SKILL.md")]:
        text = _skill_bundle(path)
        assert "sheetops doctor libreoffice" in text
        assert "sheetops recalc" in text


def test_agent_skills_document_benchmark_harness_for_operation_changes() -> None:
    for path in [Path("skills/codex/SKILL.md"), Path("skills/claude/SKILL.md")]:
        text = _skill_bundle(path)
        assert "Benchmark before optimizing." in text
        assert "python -m zalos_sheetops.benchmarks.large_xlsx" in text
        assert "src/zalos_sheetops/benchmarks/large_xlsx.py" in text
        assert "tests/test_large_xlsx_benchmark.py" in text
        assert (
            "Do not claim a performance improvement without before/after evidence."
            in text
        )
        assert "Do not commit generated benchmark artifacts." in text


def _skill_bundle(path: Path) -> str:
    entry = path.read_text()
    references = [
        path.parent / target
        for target in re.findall(r"\]\((references/[^)]+)\)", entry)
    ]
    return "\n".join([entry, *(ref.read_text() for ref in references)])


def test_skill_reference_links_resolve_from_both_installations() -> None:
    for host in ("codex", "claude"):
        path = Path("skills") / host / "SKILL.md"
        for target in re.findall(r"\]\((references/[^)]+)\)", path.read_text()):
            reference = path.parent / target
            assert reference.is_file(), str(reference)
            assert reference.resolve().is_relative_to(Path("skills").resolve())
