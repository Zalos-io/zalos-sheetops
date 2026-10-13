import json
from decimal import localcontext

import pytest

from zalos_sheetops.cli import app
from zalos_sheetops.testing import CliRunner


def run_check(**changes: object) -> tuple[int, dict]:
    request = {"left": ["0.104", "0.104"], "right": ["0.208"], "decimalPlaces": 2}
    request.update(changes)
    result = CliRunner().invoke(app, ["check-totals", "--json", json.dumps(request)])
    return result.exit_code, json.loads(result.output)


@pytest.mark.parametrize(
    ("left", "right", "difference"),
    [(["0.104", "0.104"], ["0.208"], "-0.01"), (["0.106", "0.106"], ["0.212"], "0.01")],
)
def test_round_each_line_before_summing(
    left: list, right: list, difference: str
) -> None:
    code, response = run_check(left=left, right=right)
    assert code == 1
    assert response["ok"] is False
    assert response["errorCode"] == "totals_mismatch"
    report = response["details"]
    assert report["source"]["matches"] is True
    assert report["output"]["matches"] is False
    assert report["output"]["difference"] == difference
    assert report["validated"] is False


def test_success_uses_exact_decimal_strings_independent_of_context() -> None:
    with localcontext() as context:
        context.prec = 3
        code, response = run_check(
            left=["12345678901234567890.10", "0.20"], right=["12345678901234567890.30"]
        )
    assert code == 0
    assert response["data"]["validated"] is True
    assert response["data"]["output"]["leftTotal"] == "12345678901234567890.30"


def test_source_tolerance_does_not_hide_output_difference() -> None:
    code, response = run_check(right=["0.20800000001"], sourceTolerance="0.000000001")
    assert code == 1
    assert response["details"]["source"]["matches"] is True
    assert response["details"]["output"]["difference"] == "-0.01"


def test_source_mismatch_is_not_erased_by_rounding() -> None:
    code, response = run_check(left=["1.001"], right=["1.002"])
    assert code == 1
    assert response["details"]["source"]["matches"] is False
    assert response["details"]["output"]["matches"] is True


@pytest.mark.parametrize(
    ("places", "rounding", "total"),
    [
        (2, "half_up", "-1.01"),
        (2, "half_even", "-1.00"),
        (0, "half_up", "-1"),
        (3, "half_up", "-1.005"),
    ],
)
def test_explicit_precision_and_rounding(
    places: int, rounding: str, total: str
) -> None:
    code, response = run_check(
        left=["-1.005"], right=["-1.005"], decimalPlaces=places, rounding=rounding
    )
    assert code == 0
    assert response["data"]["output"]["leftTotal"] == total


@pytest.mark.parametrize(
    "changes",
    [
        {"left": []},
        {"left": [None]},
        {"left": [True]},
        {"left": [""]},
        {"left": ["NaN"]},
        {"left": ["Infinity"]},
        {"left": ["1,000"]},
        {"left": ["1e1000000"]},
        {"left": ["1e" + "9" * 200]},
        {"left": [0.1]},
        {"decimalPlaces": True},
        {"decimalPlaces": -1},
        {"decimalPlaces": 13},
        {"rounding": "guess"},
        {"sourceTolerance": "-0.01"},
        {"sourceTolerance": "NaN"},
        {"outputTolerance": "0.01"},
    ],
)
def test_invalid_inputs_fail_closed(changes: dict) -> None:
    code, response = run_check(**changes)
    assert code == 1
    assert response["ok"] is False
    assert response["errorCode"] == "invalid_field"


def test_precision_is_required_and_command_is_discoverable() -> None:
    result = CliRunner().invoke(
        app, ["check-totals", "--json", '{"left":["1"],"right":["1"]}']
    )
    assert result.exit_code == 1
    assert json.loads(result.output)["errorCode"] == "missing_required_field"
    schema = json.loads(CliRunner().invoke(app, ["schema"]).output)
    assert any(c["name"] == "check-totals" for c in schema["data"]["commands"])


def test_command_is_listed_in_root_help() -> None:
    assert "check-totals" in CliRunner().invoke(app, ["--help"]).output
