import json
import subprocess
import sys

from zalos_sheetops.cli import app
from zalos_sheetops.testing import CliRunner

FORBIDDEN_STARTUP_MODULES = ("typer", "click", "rich", "openpyxl", "numpy", "decimal")


def test_cli_import_does_not_load_heavy_modules() -> None:
    probe = (
        "import sys, json\n"
        "import zalos_sheetops.cli\n"
        f"forbidden = {FORBIDDEN_STARTUP_MODULES!r}\n"
        "loaded = sorted(m for m in forbidden if m in sys.modules)\n"
        "print(json.dumps(loaded))\n"
    )
    result = subprocess.run(
        [sys.executable, "-c", probe],
        capture_output=True,
        text=True,
        check=True,
    )
    assert json.loads(result.stdout) == []


def test_root_help_is_plain_text_without_padding() -> None:
    result = CliRunner().invoke(app, ["--help"])

    assert result.exit_code == 0
    assert "╭" not in result.output
    assert "│" not in result.output
    for line in result.output.splitlines():
        assert line == line.rstrip()


def test_subcommand_help_is_plain_text() -> None:
    result = CliRunner().invoke(app, ["cells", "read", "--help"])

    assert result.exit_code == 0
    assert "╭" not in result.output
    assert "--json" in result.output
    assert "--raw-output" in result.output


def test_no_arguments_prints_help_and_exits_two() -> None:
    result = CliRunner().invoke(app, [])

    assert result.exit_code == 2
    assert "Agent-native Excel and CSV operations for AI agents." in result.output


def test_unknown_command_is_usage_error() -> None:
    result = CliRunner().invoke(app, ["frobnicate"])

    assert result.exit_code == 2


def test_bare_group_is_usage_error() -> None:
    result = CliRunner().invoke(app, ["cells"])

    assert result.exit_code == 2


def test_unknown_subcommand_is_usage_error() -> None:
    result = CliRunner().invoke(app, ["cells", "frobnicate"])

    assert result.exit_code == 2


def test_missing_json_option_is_usage_error() -> None:
    result = CliRunner().invoke(app, ["inspect"])

    assert result.exit_code == 2


def test_unknown_option_is_usage_error() -> None:
    result = CliRunner().invoke(app, ["inspect", "--json", "{}", "--frobnicate"])

    assert result.exit_code == 2


def test_unexpected_extra_argument_is_usage_error() -> None:
    result = CliRunner().invoke(app, ["schema", "extra"])

    assert result.exit_code == 2


def test_json_equals_form_is_accepted() -> None:
    result = CliRunner().invoke(
        app,
        ["inspect", '--json={"path":"missing-workbook.xlsx"}', "--raw-output"],
    )

    assert result.exit_code == 1
    payload = json.loads(result.output)
    assert payload["ok"] is False
    assert payload["errorCode"] == "file_not_found"
