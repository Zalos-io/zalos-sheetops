import json
from pathlib import Path

from zalos_sheetops.cli import app
from zalos_sheetops.testing import CliRunner


def _parse_output(output: str) -> dict[str, object]:
    return json.loads(output)


def test_doctor_libreoffice_reports_missing_runtime(monkeypatch) -> None:
    monkeypatch.setattr(
        "zalos_sheetops.runtime.libreoffice.detect_libreoffice",
        lambda: {
            "available": False,
            "errorCode": "libreoffice_missing",
            "executable": "",
            "platform": "linux",
            "recommendedCommand": (
                "sudo apt-get update && "
                "sudo apt-get install -y libreoffice libreoffice-calc"
            ),
        },
    )
    runner = CliRunner()

    result = runner.invoke(app, ["doctor", "libreoffice", "--raw-output"])

    assert result.exit_code == 1
    assert _parse_output(result.output) == {
        "ok": False,
        "command": "doctor libreoffice",
        "errorCode": "libreoffice_missing",
        "message": "LibreOffice executable was not found",
        "details": {
            "available": False,
            "errorCode": "libreoffice_missing",
            "executable": "",
            "platform": "linux",
            "recommendedCommand": (
                "sudo apt-get update && "
                "sudo apt-get install -y libreoffice libreoffice-calc"
            ),
        },
        "warnings": [],
        "artifacts": [],
    }


def test_doctor_libreoffice_reports_broken_runtime(monkeypatch) -> None:
    monkeypatch.setattr(
        "zalos_sheetops.runtime.libreoffice.detect_libreoffice",
        lambda: {
            "available": False,
            "errorCode": "libreoffice_broken",
            "executable": "/usr/local/bin/soffice",
            "source": "path",
            "platform": "linux",
            "stderr": "No such file or directory",
            "recommendedCommand": (
                "sudo apt-get update && "
                "sudo apt-get install -y libreoffice libreoffice-calc"
            ),
        },
    )
    runner = CliRunner()

    result = runner.invoke(app, ["doctor", "libreoffice", "--raw-output"])

    assert result.exit_code == 1
    payload = _parse_output(result.output)
    assert payload["errorCode"] == "libreoffice_broken"
    assert payload["message"] == "LibreOffice executable is broken"
    assert payload["details"]["stderr"] == "No such file or directory"


def test_doctor_libreoffice_reports_available_runtime(
    tmp_path: Path,
    monkeypatch,
) -> None:
    executable = tmp_path / "soffice"
    monkeypatch.setattr(
        "zalos_sheetops.runtime.libreoffice.detect_libreoffice",
        lambda: {
            "available": True,
            "executable": str(executable),
            "source": "path",
            "platform": "linux",
        },
    )
    runner = CliRunner()

    result = runner.invoke(app, ["doctor", "libreoffice", "--raw-output"])

    assert result.exit_code == 0
    payload = _parse_output(result.output)
    assert payload["ok"] is True
    assert payload["command"] == "doctor libreoffice"
    assert payload["data"]["executable"] == str(executable)
