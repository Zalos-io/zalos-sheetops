from decimal import Decimal, InvalidOperation
from pathlib import Path

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.xlsx_xml import (
    CellData,
    read_labeled_cells_xml,
    read_workbook_xml_summary,
)


def compare_workbook_structure(
    expected: Path,
    actual: Path,
    compare_formula_counts: bool = True,
) -> dict[str, object]:
    expected_summary = read_workbook_xml_summary(expected, max_formula_cells=0)
    actual_summary = read_workbook_xml_summary(actual, max_formula_cells=0)
    expected_sheets = _sheets_by_name(expected_summary)
    actual_sheets = _sheets_by_name(actual_summary)

    added = [
        _sheet_payload(actual_sheets[name])
        for name in actual_sheets
        if name not in expected_sheets
    ]
    removed = [
        _sheet_payload(expected_sheets[name])
        for name in expected_sheets
        if name not in actual_sheets
    ]
    changed = []
    for name in expected_sheets:
        if name not in actual_sheets:
            continue
        changes = _sheet_changes(
            expected_sheets[name],
            actual_sheets[name],
            compare_formula_counts,
        )
        if changes:
            changed.append({"name": name, "changes": changes})

    summary = {
        "added": len(added),
        "removed": len(removed),
        "changed": len(changed),
    }
    return {
        "same": summary == {"added": 0, "removed": 0, "changed": 0},
        "summary": summary,
        "addedSheets": added,
        "removedSheets": removed,
        "changedSheets": changed,
    }


def compare_labeled_cells(
    expected: Path,
    actual: Path,
    *,
    sheet_name: str,
    actual_sheet_name: str | None = None,
    label_column: str,
    value_column: str,
    labels: list[str],
    compare_formulas: bool = True,
    amount_tolerance: str = "0.00",
) -> dict[str, object]:
    if not labels:
        raise SheetOpsError(
            "invalid_field",
            "Field must not be empty",
            {"field": "labels", "expected": "non-empty list[string]"},
        )
    tolerance = _parse_amount_tolerance(amount_tolerance)
    requested_labels = set(labels)
    expected_cells = read_labeled_cells_xml(
        expected,
        sheet_name,
        label_column,
        value_column,
        requested_labels,
    )
    actual_cells = read_labeled_cells_xml(
        actual,
        actual_sheet_name or sheet_name,
        label_column,
        value_column,
        requested_labels,
    )

    comparisons = []
    matched = 0
    changed = 0
    missing = 0
    for label in labels:
        comparison = _compare_labeled_cell(
            label,
            expected_cells.get(label),
            actual_cells.get(label),
            compare_formulas,
            tolerance,
        )
        comparisons.append(comparison)
        if comparison["status"] == "matched":
            matched += 1
        elif comparison["status"] == "missing":
            missing += 1
        else:
            changed += 1

    summary = {"matched": matched, "changed": changed, "missing": missing}
    return {
        "same": changed == 0 and missing == 0,
        "summary": summary,
        "comparisons": comparisons,
    }


def _sheets_by_name(summary: dict[str, object]) -> dict[str, dict[str, object]]:
    sheets = summary["sheets"]
    if not isinstance(sheets, list):
        return {}
    result = {}
    for position, sheet in enumerate(sheets, start=1):
        if not isinstance(sheet, dict):
            continue
        name = sheet.get("name")
        if isinstance(name, str):
            result[name] = {**sheet, "position": position}
    return result


def _sheet_payload(sheet: dict[str, object]) -> dict[str, object]:
    return {
        "name": sheet["name"],
        "position": sheet["position"],
        "dimensions": sheet["dimensions"],
        "hidden": sheet["hidden"],
        "formulaCount": sheet["formulaCount"],
    }


def _sheet_changes(
    expected: dict[str, object],
    actual: dict[str, object],
    compare_formula_counts: bool,
) -> list[dict[str, object]]:
    fields = [
        "position",
        "dimensions",
        "maxRow",
        "maxColumn",
        "hidden",
    ]
    if compare_formula_counts:
        fields.append("formulaCount")
    changes = []
    for field in fields:
        expected_value = expected.get(field)
        actual_value = actual.get(field)
        if expected_value != actual_value:
            changes.append(
                {
                    "field": field,
                    "expected": expected_value,
                    "actual": actual_value,
                }
            )
    return changes


def _compare_labeled_cell(
    label: str,
    expected: CellData | None,
    actual: CellData | None,
    compare_formulas: bool,
    amount_tolerance: Decimal,
) -> dict[str, object]:
    if expected is None or actual is None:
        differences = []
        if expected is None:
            differences.append({"field": "expected", "actual": "present"})
        if actual is None:
            differences.append({"field": "actual", "expected": "present"})
        return {
            "label": label,
            "status": "missing",
            "expected": _cell_compare_payload(expected),
            "actual": _cell_compare_payload(actual),
            "differences": differences,
        }

    differences = []
    if not _values_equal(expected.value, actual.value, amount_tolerance):
        differences.append(
            {
                "field": "value",
                "expected": expected.value,
                "actual": actual.value,
            }
        )
    if compare_formulas and expected.formula != actual.formula:
        differences.append(
            {
                "field": "formula",
                "expected": expected.formula,
                "actual": actual.formula,
            }
        )
    return {
        "label": label,
        "status": "matched" if not differences else "changed",
        "expected": _cell_compare_payload(expected),
        "actual": _cell_compare_payload(actual),
        "differences": differences,
    }


def _cell_compare_payload(cell: CellData | None) -> dict[str, object] | None:
    if cell is None:
        return None
    return {
        "cell": cell.address,
        "value": cell.value,
        "formula": cell.formula,
    }


def _values_equal(expected: object, actual: object, amount_tolerance: Decimal) -> bool:
    expected_decimal = _decimal_value(expected)
    actual_decimal = _decimal_value(actual)
    if expected_decimal is not None and actual_decimal is not None:
        return abs(expected_decimal - actual_decimal) <= amount_tolerance
    return expected == actual


def _decimal_value(value: object) -> Decimal | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value))
        except InvalidOperation:
            return None
    return None


def _parse_amount_tolerance(value: str) -> Decimal:
    try:
        decimal = Decimal(value)
    except InvalidOperation as error:
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "amountTolerance", "expected": "decimal string"},
        ) from error
    if not decimal.is_finite() or decimal < 0:
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "amountTolerance", "expected": "decimal string"},
        )
    return decimal
