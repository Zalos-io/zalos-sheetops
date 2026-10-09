import json

from zalos_sheetops.cli import app
from zalos_sheetops.testing import CliRunner


def test_default_output_is_compact_single_line_json() -> None:
    result = CliRunner().invoke(app, ["schema"])

    assert result.exit_code == 0
    body = result.output.rstrip("\n")
    assert "\n" not in body
    assert '": ' not in body
    assert '", "' not in body
    assert json.loads(body)["ok"] is True


def test_raw_output_flag_still_emits_compact_json() -> None:
    result = CliRunner().invoke(app, ["schema", "--raw-output"])

    assert result.exit_code == 0
    body = result.output.rstrip("\n")
    assert "\n" not in body
    assert json.loads(body)["ok"] is True


def test_pretty_flag_emits_indented_json() -> None:
    result = CliRunner().invoke(app, ["schema", "--pretty"])

    assert result.exit_code == 0
    assert result.output.startswith("{\n")
    assert json.loads(result.output)["ok"] is True
