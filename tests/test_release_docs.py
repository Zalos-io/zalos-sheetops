from pathlib import Path


def test_ci_runs_documented_verification_commands() -> None:
    workflow = Path(".github/workflows/ci.yml").read_text()

    for command in [
        "rtk uv run pytest",
        "rtk uv run ruff check .",
        "rtk uv run ruff format --check .",
    ]:
        assert command in workflow


def test_readme_documents_public_consumption_path() -> None:
    readme = Path("README.md").read_text()

    assert "pip install zalos-sheetops==0.4.2" in readme
    assert "v0.4.0" in readme
    assert "https://github.com/Zalos-io/zalos-sheetops" in readme
