from zalos_sheetops.cli import app
from zalos_sheetops.testing import CliRunner


def _help_output(args: list[str]) -> str:
    result = CliRunner().invoke(app, args, color=False)
    assert result.exit_code == 0
    return result.output


def test_top_level_help_describes_sheetops_and_command_groups() -> None:
    output = _help_output(["--help"])

    assert "Agent-native Excel and CSV operations for AI agents." in output
    assert "Inspect workbook metadata" in output
    assert "Read bounded cell ranges" in output
    assert "Validate targeted cell hyperlinks" in output


def test_inspect_help_explains_xml_default_and_json_options() -> None:
    output = _help_output(["inspect", "--help"])

    assert "Inspect workbook metadata" in output
    assert "engine defaults to auto/XML" in output
    assert "JSON request payload" in output
    assert "compact JSON" in output


def test_cells_read_help_explains_bounded_range_contract() -> None:
    output = _help_output(["cells", "read", "--help"])

    assert "Read a bounded A1-style cell range" in output
    assert "Use engine=openpyxl only for targeted" in output
    assert "ranges." in output


def test_hyperlinks_assert_help_explains_cell_targeted_validation() -> None:
    output = _help_output(["hyperlinks", "assert", "--help"])

    assert "Validate targeted cell hyperlinks" in output
    assert "target, display, and tooltip" in output
