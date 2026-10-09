from pathlib import Path

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_table import read_table
from zalos_sheetops.normalize.amounts import normalize_amount
from zalos_sheetops.normalize.dates import normalize_date
from zalos_sheetops.normalize.text import canonical_header


def preview_normalization(
    path: Path,
    sheet_name: str | None,
    header_strategy: str,
    header_row: int | None,
    columns: list[dict[str, str]],
    max_rows: int | None,
) -> dict[str, object]:
    if max_rows is not None and max_rows < 1:
        raise SheetOpsError(
            "invalid_max_rows",
            "maxRows must be a positive integer",
            {"maxRows": max_rows},
        )
    table = _read_table_for_preview(
        path,
        sheet_name,
        header_strategy,
        header_row,
        columns,
        max_rows,
    )
    rules = _normalize_rules(columns, table["headers"])
    preview_rows = []
    issue_count = 0

    for offset, row in enumerate(
        table["rows"][: max_rows or len(table["rows"])],
        start=1,
    ):
        output_values: dict[str, object] = {}
        issues: list[dict[str, object]] = []
        for rule in rules:
            index = rule["index"]
            raw_value = row[index] if index < len(row) else ""
            try:
                output_values[rule["name"]] = _normalize_value(
                    rule["kind"],
                    raw_value,
                )
            except SheetOpsError as error:
                output_values[rule["name"]] = {
                    "raw": "" if raw_value is None else str(raw_value),
                    "normalized": "",
                }
                issues.append(
                    {
                        "column": rule["source"],
                        "kind": rule["kind"],
                        "errorCode": error.error_code,
                    }
                )
        issue_count += len(issues)
        preview_rows.append(
            {
                "rowNumber": table["headerRow"] + offset,
                "values": output_values,
                "issues": issues,
            }
        )

    return {
        "rowCount": table["rowCount"],
        "previewRows": preview_rows,
        "issueCount": issue_count,
    }


def _normalize_rules(
    columns: list[dict[str, str]],
    headers: list[str],
) -> list[dict[str, object]]:
    header_set = set(headers)
    rules: list[dict[str, object]] = []
    for column in columns:
        source = column.get("source", "")
        kind = column.get("kind", "")
        name = column.get("name", source)
        source_canonical = canonical_header(source)
        if kind not in {"amount", "date"}:
            raise SheetOpsError(
                "invalid_normalization_kind",
                "Unsupported normalization kind",
                {"kind": kind},
            )
        if source_canonical not in header_set:
            raise SheetOpsError(
                "missing_normalization_column",
                "Normalization column does not exist",
                {"source": source},
            )
        rules.append(
            {
                "source": source,
                "index": headers.index(source_canonical),
                "kind": kind,
                "name": name,
            }
        )
    return rules


def _read_table_for_preview(
    path: Path,
    sheet_name: str | None,
    header_strategy: str,
    header_row: int | None,
    columns: list[dict[str, str]],
    max_rows: int | None,
) -> dict[str, object]:
    try:
        return read_table(path, sheet_name, header_strategy, header_row, max_rows)
    except SheetOpsError as error:
        if error.error_code != "ambiguous_header_row" or header_strategy != "detect":
            raise
        source_headers = {
            canonical_header(column.get("source", "")) for column in columns
        }
        candidate_rows = error.details.get("headerRows", [])
        matches = []
        for candidate_row in candidate_rows:
            table = read_table(path, sheet_name, "row", candidate_row, max_rows)
            if source_headers.issubset(set(table["headers"])):
                matches.append(table)
        if len(matches) == 1:
            return matches[0]
        raise


def _normalize_value(kind: str, value: object) -> dict[str, object]:
    if kind == "amount":
        return normalize_amount(value)
    if kind == "date":
        return normalize_date(value)
    raise SheetOpsError(
        "invalid_normalization_kind",
        "Unsupported normalization kind",
        {"kind": kind},
    )
