"""Fresh grading computes formulas instead of trusting saved results."""

import json
import shutil
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path
from zipfile import ZipFile

import pytest
from openpyxl import Workbook

from zalos_sheetops.benchmarks.public import regrade
from zalos_sheetops.benchmarks.public.fixtures import file_sha256
from zalos_sheetops.benchmarks.public.grading import grade_workbook

SHEET = "xl/worksheets/sheet1.xml"


def _books(folder: Path, formula: str, saved: bytes) -> list[Path]:
    """Input A1=2, A2=3; golden A3 =A1+A2 saved as 5; output A3 saved as given."""
    folder.mkdir(parents=True, exist_ok=True)
    paths = [folder / name for name in ("initial.xlsx", "golden.xlsx", "output.xlsx")]
    book = Workbook()
    book.active["A1"], book.active["A2"] = 2, 3
    book.save(paths[0])
    for path, cell, result in [(paths[1], "=A1+A2", b"5"), (paths[2], formula, saved)]:
        book.active["A3"] = cell
        book.save(path)
        with ZipFile(path) as archive:
            parts = {item: archive.read(item) for item in archive.infolist()}
        with ZipFile(path, "w") as archive:
            for item, data in parts.items():
                if item.filename == SHEET:
                    data = data.replace(b"<v />", b"<v>" + result + b"</v>")
                archive.writestr(item, data)
    return paths


def _engine(seen: list[bytes]) -> Callable[..., subprocess.CompletedProcess]:
    """Stand-in for LibreOffice: record the sheet it receives and copy it out."""

    def run(command: list[str], **kwargs: object) -> subprocess.CompletedProcess:
        staged = Path(command[-1])
        with ZipFile(staged) as archive:
            seen.append(archive.read(SHEET))
        (Path(command[-2]) / staged.name).write_bytes(staged.read_bytes())
        return subprocess.CompletedProcess(command, 0, "", "")

    return run


@pytest.mark.parametrize(
    ("fresh", "engine_sees"),
    [(False, b"<f>A1+A2</f><v>99</v>"), (True, b"<f>A1+A2</f></c>")],
)
def test_fresh_grade_clears_saved_formula_results_before_the_engine(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, fresh: bool, engine_sees: bytes
) -> None:
    initial, golden, output = _books(tmp_path, "=A1+A2", b"99")
    sources = {path: path.read_bytes() for path in (golden, output)}
    seen: list[bytes] = []
    monkeypatch.setattr(subprocess, "run", _engine(seen))
    grade_workbook(
        initial,
        golden,
        output,
        {"answer_position": "A3"},
        Path("soffice"),
        tmp_path / "evaluation",
        fresh_formula_results=fresh,
    )
    assert len(seen) == 2
    assert engine_sees in seen[1]
    assert all(path.read_bytes() == data for path, data in sources.items())


def test_fresh_regrade_makes_its_own_evaluation_copies(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    dataset, runs = tmp_path / "data", tmp_path / "agents"
    attempt = runs / "runs/1/baseline/0"
    initial, golden, output = _books(dataset, "=A1+A2", b"99")
    (attempt / "workspace").mkdir(parents=True)
    output = output.rename(attempt / "workspace/output.xlsx")
    # Copies from an earlier default regrade still hold the saved results.
    (attempt / "evaluation/golden").mkdir(parents=True)
    (attempt / "evaluation/actual").mkdir()
    shutil.copyfile(golden, attempt / "evaluation/golden" / golden.name)
    shutil.copyfile(output, attempt / "evaluation/actual/output.xlsx")
    record = {
        "task": 1,
        "arm": "baseline",
        "repeat": 0,
        "grading": {"status": "incorrect"},
        "usage": {},
        "inputSha256": file_sha256(initial),
        "goldenSha256": file_sha256(golden),
        "outputSha256": file_sha256(output),
    }
    (runs / "agents.jsonl").write_text(json.dumps(record) + "\n")
    tasks = tmp_path / "tasks.json"
    tasks.write_text(
        json.dumps([{"id": 1, "spreadsheet_path": ".", "answer_position": "A3"}])
    )
    seen: list[bytes] = []
    monkeypatch.setattr(subprocess, "run", _engine(seen))
    command = [
        "regrade",
        *["--dataset", str(dataset), "--tasks", str(tasks), "--runs", str(runs)],
        *["--soffice", "soffice"],
    ]
    monkeypatch.setattr(sys, "argv", [*command, "--output", str(tmp_path / "a.jsonl")])
    regrade.main()
    assert seen == []
    fresh = [*command, "--output", str(tmp_path / "b.jsonl"), "--fresh-formula-results"]
    monkeypatch.setattr(sys, "argv", fresh)
    regrade.main()
    assert len(seen) == 2
    assert all(b"<v>" not in xml.split(b'r="A3"', 1)[1] for xml in seen)
    assert (attempt / "evaluation-fresh/actual/output.xlsx").exists()
    default = json.loads((tmp_path / "a.jsonl").read_text())
    assert "formulaEvaluation" not in default["evaluationEngine"]
    graded = json.loads((tmp_path / "b.jsonl").read_text())
    assert graded["evaluationEngine"]["formulaEvaluation"] == "fresh"


def test_fresh_formula_results_fail_closed_without_libreoffice(tmp_path: Path) -> None:
    initial, golden, output = _books(tmp_path, "=A1+A2", b"99")
    with pytest.raises(ValueError, match="soffice"):
        grade_workbook(
            initial,
            golden,
            output,
            {"answer_position": "A3"},
            fresh_formula_results=True,
        )
    with pytest.raises(ValueError, match="native"):
        regrade.regrade_records(
            tmp_path,
            [],
            tmp_path,
            tmp_path / "graded.jsonl",
            native_evaluation_root=tmp_path,
            fresh_formula_results=True,
        )


@pytest.mark.parametrize(
    ("formula", "saved", "status"),
    [("=A1+A2", b"99", "pass"), ("=A1*A2", b"5", "incorrect")],
)
def test_real_fresh_grade_scores_formulas_not_saved_results(
    tmp_path: Path, formula: str, saved: bytes, status: str
) -> None:
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable is None:
        pytest.skip("LibreOffice is not available on this test host")
    initial, golden, output = _books(tmp_path, formula, saved)
    result = grade_workbook(
        initial,
        golden,
        output,
        {"answer_position": "A3"},
        Path(executable),
        tmp_path / "evaluation",
        fresh_formula_results=True,
    )
    assert result["status"] == status
