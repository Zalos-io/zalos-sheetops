import shutil
import subprocess
import tempfile
from pathlib import Path

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.formulas.scan import scan_formulas
from zalos_sheetops.runtime.libreoffice import detect_libreoffice


def recalc_workbook(
    path: Path,
    output_dir: Path,
    filename: str | None,
    executable: str | None = None,
) -> dict[str, object]:
    output_name = (
        _validate_output_filename(filename)
        if filename is not None
        else f"{path.stem}-recalculated{path.suffix}"
    )
    output_path = output_dir / output_name
    if output_path.resolve(strict=False) == path.resolve(strict=False):
        raise SheetOpsError(
            "invalid_path",
            "Output filename must not overwrite the source workbook",
            {"path": str(output_path), "sourcePath": str(path)},
        )

    detected = detect_libreoffice(executable)
    if not detected["available"]:
        error_code = str(detected.get("errorCode", "libreoffice_missing"))
        recalc_error_code = (
            "recalc_executable_broken"
            if error_code == "libreoffice_broken"
            else "recalc_executable_missing"
        )
        message = (
            "LibreOffice executable is broken"
            if error_code == "libreoffice_broken"
            else "LibreOffice executable was not found"
        )
        raise SheetOpsError(
            recalc_error_code,
            message,
            detected,
        )
    office = str(detected["executable"])
    output_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as staging_dir_name:
        staging_dir = Path(staging_dir_name)
        command = [
            office,
            "--headless",
            "--convert-to",
            "xlsx",
            str(path),
            "--outdir",
            str(staging_dir),
        ]
        try:
            completed = subprocess.run(
                command,
                check=False,
                capture_output=True,
                text=True,
            )
        except OSError as error:
            raise SheetOpsError(
                "recalc_failed",
                "LibreOffice recalculation failed",
                {"error": str(error)},
            ) from error
        if completed.returncode != 0:
            raise SheetOpsError(
                "recalc_failed",
                "LibreOffice recalculation failed",
                {"stderr": completed.stderr.strip()},
            )
        converted_path = staging_dir / path.name
        if not converted_path.exists():
            raise SheetOpsError(
                "recalc_failed",
                "LibreOffice did not create expected output",
                {"path": str(converted_path)},
            )
        shutil.copy2(converted_path, output_path)
    if not output_path.exists():
        raise SheetOpsError(
            "recalc_failed",
            "LibreOffice did not create expected output",
            {"path": str(output_path)},
        )

    scan = scan_formulas(output_path)
    return {
        "sourcePath": str(path),
        "recalculatedPath": str(output_path),
        "formulaCount": scan["formulaCount"],
        "errorTokenCount": scan["errorTokenCount"],
        "errorTokens": scan["errorTokens"],
    }


def _validate_output_filename(filename: str) -> str:
    candidate = Path(filename)
    is_plain_filename = (
        candidate.name
        and candidate.name == filename
        and candidate.name not in {".", ".."}
    )
    if not is_plain_filename:
        raise SheetOpsError(
            "invalid_path",
            "Output filename must be a plain filename",
            {"path": filename},
        )
    return candidate.name
