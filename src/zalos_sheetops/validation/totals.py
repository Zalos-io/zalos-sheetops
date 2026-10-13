"""Compare complete amount populations before and after per-value rounding."""

import re
from decimal import (
    ROUND_HALF_EVEN,
    ROUND_HALF_UP,
    Context,
    Decimal,
    InvalidOperation,
    localcontext,
)

from zalos_sheetops.errors import SheetOpsError

_ROUNDING = {"half_up": ROUND_HALF_UP, "half_even": ROUND_HALF_EVEN}
_DECIMAL = re.compile(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?\Z")


def _invalid(field: str, expected: str) -> SheetOpsError:
    return SheetOpsError(
        "invalid_field",
        "Invalid totals check input",
        {
            "field": field,
            "expected": expected,
        },
    )


def _decimal(value: object, field: str) -> Decimal:
    # Strings retain decimal source precision across JSON and Python boundaries.
    if not isinstance(value, str) or len(value) > 256 or not _DECIMAL.fullmatch(value):
        raise _invalid(field, "finite decimal string (no separators)")
    try:
        amount = Decimal(value, context=Context())
    except InvalidOperation as error:
        raise _invalid(field, "bounded finite decimal string") from error
    if abs(amount.as_tuple().exponent) > 128 or len(amount.as_tuple().digits) > 128:
        raise _invalid(field, "at most 128 digits and exponent magnitude 128")
    return amount


def _amounts(values: object, field: str) -> list[Decimal]:
    if not isinstance(values, list) or not 1 <= len(values) <= 100000:
        raise _invalid(field, "non-empty list of at most 100000 decimal strings")
    return [_decimal(value, f"{field}[{i}]") for i, value in enumerate(values)]


def _comparison(left: Decimal, right: Decimal, tolerance: Decimal) -> dict[str, object]:
    difference = left - right
    return {
        "leftTotal": format(left.copy_abs() if left == 0 else left, "f"),
        "rightTotal": format(right.copy_abs() if right == 0 else right, "f"),
        "difference": format(
            difference.copy_abs() if difference == 0 else difference, "f"
        ),
        "matches": abs(difference) <= tolerance,
    }


def check_totals(
    left: list[str],
    right: list[str],
    decimal_places: int,
    rounding: str = "half_up",
    source_tolerance: str = "0",
) -> dict[str, object]:
    """Check source sums and sums of rounded lines; never alter caller values.

    Pass complete populations of comparable units, and run separately per group.
    This is arithmetic validation, not formula evaluation or completeness proof.
    A source tolerance never permits a mismatch at output precision.
    """
    if type(decimal_places) is not int or not 0 <= decimal_places <= 12:
        raise _invalid("decimalPlaces", "integer from 0 to 12")
    if not isinstance(rounding, str) or rounding not in _ROUNDING:
        raise _invalid("rounding", "half_up | half_even")
    lhs = _amounts(left, "left")
    rhs = _amounts(right, "right")
    tolerance = _decimal(source_tolerance, "sourceTolerance")
    if tolerance < 0:
        raise _invalid("sourceTolerance", "non-negative decimal string")
    # Covers the bounded input exponents, digits and population carry, independent
    # of callers' Decimal precision, traps and rounding configuration.
    with localcontext(Context(prec=400)):
        quantum = Decimal(1).scaleb(-decimal_places)
        source = _comparison(sum(lhs, Decimal(0)), sum(rhs, Decimal(0)), tolerance)
        mode = _ROUNDING[rounding]
        output = _comparison(
            sum((v.quantize(quantum, rounding=mode) for v in lhs), Decimal(0)),
            sum((v.quantize(quantum, rounding=mode) for v in rhs), Decimal(0)),
            Decimal(0),
        )
    return {
        "validated": source["matches"] and output["matches"],
        "decimalPlaces": decimal_places,
        "rounding": rounding,
        "sourceTolerance": source_tolerance,
        "leftCount": len(lhs),
        "rightCount": len(rhs),
        "source": source,
        "output": output,
    }


def check_totals_request(payload: dict[str, object]) -> dict[str, object]:
    allowed = {"left", "right", "decimalPlaces", "rounding", "sourceTolerance"}
    for field in sorted(payload.keys() - allowed):
        raise _invalid(field, "supported check-totals field")
    for field in ("left", "right", "decimalPlaces"):
        if field not in payload:
            raise SheetOpsError(
                "missing_required_field", "Required field is missing", {"field": field}
            )
    return check_totals(
        payload["left"],
        payload["right"],
        payload["decimalPlaces"],
        payload.get("rounding", "half_up"),
        payload.get("sourceTolerance", "0"),
    )
