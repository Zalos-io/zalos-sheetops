"""Zero-dependency command dispatcher.

The CLI is invoked once per agent action, so interpreter startup dominates
latency. This module therefore avoids typer/click/rich and defers every
heavy import (openpyxl, decimal, command implementations) until a command
actually needs it.
"""

import json
import sys

from zalos_sheetops.command_registry import (
    command_manifest,
    error_response,
    success_response,
)
from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.paths import validate_agent_path

APP_HELP = "Agent-native Excel and CSV operations for AI agents."
JSON_OPTION_HELP = "JSON request payload matching the command contract."
PRETTY_OPTION_HELP = "Pretty-print the JSON response for humans."
RAW_OUTPUT_OPTION_HELP = "Deprecated alias for the default compact JSON output."


def _run_schema(json_payload: str | None, pretty: bool) -> int:
    _emit(success_response("schema", command_manifest()), pretty)
    return 0


def _run_doctor_libreoffice(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.runtime.libreoffice import detect_libreoffice

    command = "doctor libreoffice"
    result = detect_libreoffice()
    if result["available"]:
        _emit(success_response(command, result), pretty)
        return 0
    message = _libreoffice_error_message(str(result["errorCode"]))
    _emit(
        error_response(
            command,
            str(result["errorCode"]),
            message,
            result,
        ),
        pretty,
    )
    return 1


def _run_inspect(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.io.inspect import inspect_workbook

    command = "inspect"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        path = validate_agent_path(path_value)
        engine = _optional_engine_field(payload, default="auto")
        max_formula_cells = _optional_positive_integer_field(
            payload,
            "maxFormulaCells",
        )
        max_merged_ranges = _optional_positive_integer_field(
            payload,
            "maxMergedRanges",
        )
        max_graphics = _optional_positive_integer_field(
            payload,
            "maxGraphics",
        )
        max_defined_names = _optional_positive_integer_field(
            payload,
            "maxDefinedNames",
        )
        _emit(
            success_response(
                command,
                inspect_workbook(
                    path,
                    engine=engine,
                    max_formula_cells=max_formula_cells or 20,
                    max_merged_ranges=max_merged_ranges or 1000,
                    max_graphics=max_graphics or 1000,
                    max_defined_names=max_defined_names or 1000,
                ),
            ),
            pretty,
        )
        return 0
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_read_table(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.io.read_table import read_table

    command = "read-table"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        sheet = payload.get("sheet")
        if sheet is not None and not isinstance(sheet, str):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "sheet", "expected": "string | null"},
            )
        header_strategy = payload.get("headerStrategy", "detect")
        if not isinstance(header_strategy, str):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "headerStrategy", "expected": "string"},
            )
        header_row = payload.get("headerRow")
        if header_row is not None and (
            isinstance(header_row, bool) or not isinstance(header_row, int)
        ):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "headerRow", "expected": "integer | null"},
            )
        max_rows = payload.get("maxRows")
        if max_rows is not None and (
            isinstance(max_rows, bool) or not isinstance(max_rows, int) or max_rows < 1
        ):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "maxRows", "expected": "positive integer | null"},
            )
        path = validate_agent_path(path_value)
        _emit(
            success_response(
                command,
                read_table(path, sheet, header_strategy, header_row, max_rows),
            ),
            pretty,
        )
        return 0
    except (SheetOpsError, ValueError) as error:
        if isinstance(error, SheetOpsError):
            response = error_response(
                command,
                error.error_code,
                error.message,
                error.details,
            )
        else:
            response = error_response(command, "invalid_request", str(error), {})
        _emit(response, pretty)
        return 1


def _run_cells_read(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.io.read_cells import read_cells

    command = "cells read"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        sheet = _require_string_field(payload, "sheet")
        cell_range = _require_string_field(payload, "range")
        include = _optional_include_field(payload)
        engine = _optional_engine_field(payload, default="openpyxl")
        _emit(
            success_response(
                command,
                read_cells(
                    validate_agent_path(path_value),
                    sheet,
                    cell_range,
                    include,
                    engine=engine,
                ),
            ),
            pretty,
        )
        return 0
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_cells_search(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.io.search_cells import search_cells

    command = "cells search"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        query = _require_string_field(payload, "query")
        fields = _optional_non_empty_string_list_field(payload, "fields")
        regex = _optional_bool_field(payload, "regex", False)
        case_sensitive = _optional_bool_field(payload, "caseSensitive", False)
        include_hidden = _optional_bool_field(payload, "includeHidden", False)
        limit = _optional_positive_integer_field(payload, "limit") or 100
        offset = _optional_non_negative_integer_field(payload, "offset") or 0
        max_cells = _optional_positive_integer_field(payload, "maxCells") or 1_000_000
        result = search_cells(
            validate_agent_path(path_value),
            query,
            fields=fields,
            sheet_name=payload.get("sheet"),
            cell_range=payload.get("range"),
            regex=regex,
            case_sensitive=case_sensitive,
            include_hidden=include_hidden,
            limit=limit,
            offset=offset,
            max_cells=max_cells,
        )
        _emit(success_response(command, result), pretty)
        return 0
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_styles_assert(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.styles.assertions import assert_styles

    command = "styles assert"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        sheet = _require_string_field(payload, "sheet")
        checks = _require_style_checks(payload)
        result = assert_styles(validate_agent_path(path_value), sheet, checks)
        if result["valid"]:
            _emit(success_response(command, result), pretty)
            return 0
        _emit(
            error_response(
                command,
                "style_assertion_failed",
                "Style assertions failed",
                result,
            ),
            pretty,
        )
        return 1
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_hyperlinks_assert(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.hyperlinks.assertions import assert_hyperlinks

    command = "hyperlinks assert"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        sheet = _require_string_field(payload, "sheet")
        checks = _require_hyperlink_checks(payload)
        result = assert_hyperlinks(validate_agent_path(path_value), sheet, checks)
        if result["valid"]:
            _emit(success_response(command, result), pretty)
            return 0
        _emit(
            error_response(
                command,
                "hyperlink_assertion_failed",
                "Hyperlink assertions failed",
                result,
            ),
            pretty,
        )
        return 1
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_styles_read(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.styles.read import read_styles

    command = "styles read"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        sheet = _require_string_field(payload, "sheet")
        cell_range = _require_string_field(payload, "range")
        max_cells = _optional_positive_integer_field(payload, "maxCells")
        result = read_styles(
            validate_agent_path(path_value),
            sheet,
            cell_range,
            max_cells=max_cells or 10000,
        )
        _emit(success_response(command, result), pretty)
        return 0
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_validate(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.validation.workbook import (
        validate_workbook,
        validate_workbook_rules,
    )

    command = "validate"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        path = validate_agent_path(path_value)
        if "sheets" in payload:
            result = validate_workbook_rules(path, _require_sheet_rules(payload))
        else:
            required_sheets = _require_string_list_field(payload, "requiredSheets")
            required_columns = _require_string_list_mapping_field(
                payload,
                "requiredColumns",
            )
            unique_keys = _require_string_list_mapping_field(payload, "uniqueKeys")
            result = validate_workbook(
                path,
                required_sheets,
                required_columns,
                unique_keys,
            )
        if result["valid"]:
            _emit(success_response(command, result), pretty)
            return 0
        _emit(
            error_response(
                command,
                "validation_failed",
                "Workbook validation failed",
                result,
            ),
            pretty,
        )
        return 1
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_formulas_scan(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.formulas.scan import scan_formulas

    command = "formulas scan"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        sheet = payload.get("sheet")
        if sheet is not None and not isinstance(sheet, str):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "sheet", "expected": "string | null"},
            )
        cell_range = payload.get("range")
        if cell_range is not None and not isinstance(cell_range, str):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "range", "expected": "string | null"},
            )
        max_results = payload.get("maxResults")
        if max_results is not None and (
            isinstance(max_results, bool)
            or not isinstance(max_results, int)
            or max_results < 1
        ):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "maxResults", "expected": "positive integer | null"},
            )
        engine = payload.get("engine", "auto")
        if not isinstance(engine, str) or engine not in {"auto", "xml", "openpyxl"}:
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "engine", "expected": "auto | xml | openpyxl"},
            )
        include = _optional_formula_include_field(payload)
        fail_fast = payload.get("failFast", False)
        if not isinstance(fail_fast, bool):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "failFast", "expected": "boolean"},
            )
        result = scan_formulas(
            validate_agent_path(path_value),
            sheet_name=sheet,
            cell_range=cell_range,
            max_results=max_results,
            engine=engine,
            include=include,
            fail_fast=fail_fast,
        )
        if result["errorTokenCount"] == 0:
            _emit(success_response(command, result), pretty)
            return 0
        _emit(
            error_response(
                command,
                "formula_errors_found",
                "Formula error tokens were found",
                result,
            ),
            pretty,
        )
        return 1
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_display_scan(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.display.scan import scan_display_overflow

    command = "display scan"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        sheet = payload.get("sheet")
        if sheet is not None and not isinstance(sheet, str):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "sheet", "expected": "string | null"},
            )
        cell_range = payload.get("range")
        if cell_range is not None and not isinstance(cell_range, str):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "range", "expected": "string | null"},
            )
        max_results = _optional_positive_integer_field(payload, "maxResults")
        max_cells = _optional_positive_integer_field(payload, "maxCells")
        include_hidden = payload.get("includeHidden", False)
        if not isinstance(include_hidden, bool):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "includeHidden", "expected": "boolean"},
            )
        result = scan_display_overflow(
            validate_agent_path(path_value),
            sheet_name=sheet,
            cell_range=cell_range,
            max_results=max_results,
            max_cells=max_cells,
            include_hidden=include_hidden,
        )
        if result["candidateCount"] == 0:
            _emit(success_response(command, result), pretty)
            return 0
        _emit(
            error_response(
                command,
                "display_overflow_candidates_found",
                (
                    "Cells may display as hashes because their rendered value "
                    "exceeds column width"
                ),
                result,
            ),
            pretty,
        )
        return 1
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_compare_structure(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.compare.workbooks import compare_workbook_structure

    command = "compare structure"
    try:
        payload = _parse_json_payload(command, json_payload)
        expected = _require_workbook_source_field(payload, "expected")
        actual = _require_workbook_source_field(payload, "actual")
        compare_formula_counts = _optional_bool_field(
            payload,
            "compareFormulaCounts",
            True,
        )
        result = compare_workbook_structure(
            validate_agent_path(expected["path"]),
            validate_agent_path(actual["path"]),
            compare_formula_counts=compare_formula_counts,
        )
        if result["same"]:
            _emit(success_response(command, result), pretty)
            return 0
        _emit(
            error_response(
                command,
                "workbook_structure_differences_found",
                "Workbook structure differences were found",
                result,
            ),
            pretty,
        )
        return 1
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_compare_labeled_cells(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.compare.workbooks import compare_labeled_cells

    command = "compare labeled-cells"
    try:
        payload = _parse_json_payload(command, json_payload)
        expected = _require_workbook_sheet_source_field(payload, "expected")
        actual = _require_workbook_sheet_source_field(payload, "actual")
        label_column = _require_string_field(payload, "labelColumn")
        value_column = _require_string_field(payload, "valueColumn")
        labels = _require_non_empty_string_list_field(payload, "labels")
        compare_formulas = _optional_bool_field(payload, "compareFormulas", True)
        amount_tolerance = _require_amount_tolerance(payload)
        result = compare_labeled_cells(
            validate_agent_path(expected["path"]),
            validate_agent_path(actual["path"]),
            sheet_name=str(expected["sheet"]),
            actual_sheet_name=str(actual["sheet"]),
            label_column=label_column,
            value_column=value_column,
            labels=labels,
            compare_formulas=compare_formulas,
            amount_tolerance=amount_tolerance,
        )
        if result["same"]:
            _emit(success_response(command, result), pretty)
            return 0
        _emit(
            error_response(
                command,
                "labeled_cell_differences_found",
                "Labeled cell differences were found",
                result,
            ),
            pretty,
        )
        return 1
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_normalize_preview(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.normalize.preview import preview_normalization

    command = "normalize preview"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        columns = payload.get("columns")
        if not isinstance(columns, list) or any(
            not isinstance(item, dict) for item in columns
        ):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "columns", "expected": "list[columnRule]"},
            )
        sheet = payload.get("sheet")
        if sheet is not None and not isinstance(sheet, str):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "sheet", "expected": "string | null"},
            )
        header_strategy = payload.get("headerStrategy", "detect")
        header_row = payload.get("headerRow")
        if header_row is not None and (
            isinstance(header_row, bool) or not isinstance(header_row, int)
        ):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "headerRow", "expected": "integer | null"},
            )
        max_rows = payload.get("maxRows")
        if max_rows is not None and (
            isinstance(max_rows, bool) or not isinstance(max_rows, int)
        ):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "maxRows", "expected": "integer | null"},
            )
        result = preview_normalization(
            validate_agent_path(path_value),
            sheet,
            header_strategy,
            header_row,
            columns,
            max_rows,
        )
        _emit(success_response(command, result), pretty)
        return 0
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_recalc(json_payload: str | None, pretty: bool) -> int:
    from pathlib import Path

    from zalos_sheetops.recalc.libreoffice import recalc_workbook

    command = "recalc"
    try:
        payload = _parse_json_payload(command, json_payload)
        path_value = payload.get("path")
        if not isinstance(path_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "path"},
            )
        output_dir_value = payload.get("outputDir")
        if not isinstance(output_dir_value, str):
            raise SheetOpsError(
                "missing_required_field",
                "Required field is missing or invalid",
                {"field": "outputDir"},
            )
        filename = payload.get("filename")
        if filename is not None and not isinstance(filename, str):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "filename", "expected": "string | null"},
            )
        executable = payload.get("executable")
        if executable is not None and not isinstance(executable, str):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": "executable", "expected": "string | null"},
            )
        result = recalc_workbook(
            validate_agent_path(path_value),
            Path(output_dir_value).expanduser(),
            filename,
            executable,
        )
        if result["errorTokenCount"]:
            _emit(
                error_response(
                    command,
                    "formula_errors_found",
                    "Formula error tokens were found",
                    result,
                ),
                pretty,
            )
            return 1
        _emit(
            success_response(
                command,
                result,
                artifacts=[{"kind": "xlsx", "path": result["recalculatedPath"]}],
            ),
            pretty,
        )
        return 0
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_diff(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.diff.tabular import diff_tables

    command = "diff"
    try:
        payload = _parse_json_payload(command, json_payload)
        left_source = _require_table_source_field(payload, "left")
        right_source = _require_table_source_field(payload, "right")
        key_columns = _require_non_empty_string_list_field(payload, "keyColumns")
        compare_columns = _optional_non_empty_string_list_field(
            payload,
            "compareColumns",
        )
        amount_tolerance = _require_amount_tolerance(payload)
        result = diff_tables(
            _load_table_source(left_source),
            _load_table_source(right_source),
            key_columns,
            compare_columns,
            amount_tolerance,
        )
        if result["same"]:
            _emit(success_response(command, result), pretty)
            return 0
        _emit(
            error_response(
                command,
                "differences_found",
                "Tabular differences were found",
                result,
            ),
            pretty,
        )
        return 1
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _run_reconcile(json_payload: str | None, pretty: bool) -> int:
    from zalos_sheetops.reconcile.exact import reconcile_exact

    command = "reconcile"
    try:
        payload = _parse_json_payload(command, json_payload)
        left_source = _require_table_source_field(payload, "left")
        right_source = _require_table_source_field(payload, "right")
        left_key_columns = _require_non_empty_string_list_field(
            payload,
            "leftKeyColumns",
        )
        right_key_columns = _require_non_empty_string_list_field(
            payload,
            "rightKeyColumns",
        )
        left_amount_column = _require_string_field(payload, "leftAmountColumn")
        right_amount_column = _require_string_field(payload, "rightAmountColumn")
        amount_tolerance = _require_amount_tolerance(payload)
        result = reconcile_exact(
            _load_table_source(left_source),
            _load_table_source(right_source),
            left_key_columns,
            right_key_columns,
            left_amount_column,
            right_amount_column,
            amount_tolerance,
        )
        _emit(success_response(command, result), pretty)
        return 0
    except SheetOpsError as error:
        _emit(
            error_response(command, error.error_code, error.message, error.details),
            pretty,
        )
        return 1


def _parse_json_payload(command: str, json_payload: str | None) -> dict[str, object]:
    try:
        payload = json.loads(json_payload if json_payload is not None else "")
    except json.JSONDecodeError as error:
        raise SheetOpsError(
            "invalid_json",
            "Request payload is not valid JSON",
            {"position": error.pos},
        ) from error

    if not isinstance(payload, dict):
        raise SheetOpsError(
            "invalid_json",
            "Request payload must be a JSON object",
            {"command": command},
        )
    return payload


def _libreoffice_error_message(error_code: str) -> str:
    if error_code == "libreoffice_broken":
        return "LibreOffice executable is broken"
    return "LibreOffice executable was not found"


def _require_string_list_field(
    payload: dict[str, object],
    field_name: str,
) -> list[str]:
    if field_name not in payload:
        raise SheetOpsError(
            "missing_required_field",
            "Required field is missing or invalid",
            {"field": field_name},
        )
    value = payload[field_name]
    if not isinstance(value, list) or any(not isinstance(item, str) for item in value):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "list[string]"},
        )
    return value


def _require_sheet_rules(payload: dict[str, object]) -> list[dict[str, object]]:
    value = payload.get("sheets")
    if not isinstance(value, list) or any(not isinstance(item, dict) for item in value):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "sheets", "expected": "list[sheetValidationRule]"},
        )
    return value


def _require_style_checks(payload: dict[str, object]) -> list[dict[str, object]]:
    value = payload.get("checks")
    if (
        not isinstance(value, list)
        or not value
        or any(not isinstance(item, dict) for item in value)
    ):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "checks", "expected": "non-empty list[styleCheck]"},
        )
    return value


def _require_hyperlink_checks(payload: dict[str, object]) -> list[dict[str, object]]:
    value = payload.get("checks")
    if (
        not isinstance(value, list)
        or not value
        or any(not isinstance(item, dict) for item in value)
    ):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "checks", "expected": "non-empty list[hyperlinkCheck]"},
        )
    return value


def _optional_string_list_field(
    payload: dict[str, object],
    field_name: str,
) -> list[str] | None:
    value = payload.get(field_name)
    if value is None:
        return None
    if not isinstance(value, list) or any(not isinstance(item, str) for item in value):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "list[string] | null"},
        )
    return value


def _optional_include_field(payload: dict[str, object]) -> list[str] | None:
    if "include" not in payload:
        return None
    value = payload["include"]
    if not isinstance(value, list) or any(not isinstance(item, str) for item in value):
        from zalos_sheetops.io.read_cells import INCLUDE_EXPECTED

        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "include", "expected": INCLUDE_EXPECTED},
        )
    return value


def _optional_formula_include_field(payload: dict[str, object]) -> list[str] | None:
    if "include" not in payload:
        return None
    value = payload["include"]
    if (
        not isinstance(value, list)
        or not value
        or any(
            not isinstance(item, str) or item not in {"formulas", "errors"}
            for item in value
        )
    ):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "include", "expected": "non-empty list[formulas | errors]"},
        )
    return value


def _optional_positive_integer_field(
    payload: dict[str, object],
    field_name: str,
) -> int | None:
    value = payload.get(field_name)
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or value < 1:
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "positive integer | null"},
        )
    return value


def _optional_non_negative_integer_field(
    payload: dict[str, object],
    field_name: str,
) -> int | None:
    value = payload.get(field_name)
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "non-negative integer | null"},
        )
    return value


def _optional_bool_field(
    payload: dict[str, object],
    field_name: str,
    default: bool,
) -> bool:
    value = payload.get(field_name, default)
    if not isinstance(value, bool):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "boolean"},
        )
    return value


def _optional_engine_field(payload: dict[str, object], default: str) -> str:
    value = payload.get("engine", default)
    if not isinstance(value, str) or value not in {"auto", "xml", "openpyxl"}:
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "engine", "expected": "auto | xml | openpyxl"},
        )
    return value


def _require_non_empty_string_list_field(
    payload: dict[str, object],
    field_name: str,
) -> list[str]:
    value = _require_string_list_field(payload, field_name)
    if value:
        return value
    raise SheetOpsError(
        "invalid_field",
        "Field must not be empty",
        {"field": field_name, "expected": "non-empty list[string]"},
    )


def _optional_non_empty_string_list_field(
    payload: dict[str, object],
    field_name: str,
) -> list[str] | None:
    value = _optional_string_list_field(payload, field_name)
    if value is None or value:
        return value
    raise SheetOpsError(
        "invalid_field",
        "Field must not be empty",
        {"field": field_name, "expected": "non-empty list[string] | null"},
    )


def _require_string_field(payload: dict[str, object], field_name: str) -> str:
    if field_name not in payload:
        raise SheetOpsError(
            "missing_required_field",
            "Required field is missing or invalid",
            {"field": field_name},
        )
    value = payload[field_name]
    if not isinstance(value, str):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "string"},
        )
    return value


def _require_amount_tolerance(payload: dict[str, object]) -> str:
    from decimal import Decimal, InvalidOperation

    value = payload.get("amountTolerance", "0.00")
    if not isinstance(value, str):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "amountTolerance", "expected": "decimal string"},
        )
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
    return value


def _require_table_source_field(
    payload: dict[str, object],
    field_name: str,
) -> dict[str, object]:
    if field_name not in payload:
        raise SheetOpsError(
            "missing_required_field",
            "Required field is missing or invalid",
            {"field": field_name},
        )
    source = payload[field_name]
    if not isinstance(source, dict):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "tableSource"},
        )
    path = source.get("path")
    if not isinstance(path, str):
        raise SheetOpsError(
            "missing_required_field",
            "Required field is missing or invalid",
            {"field": f"{field_name}.path"},
        )
    sheet = source.get("sheet")
    if sheet is not None and not isinstance(sheet, str):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": f"{field_name}.sheet", "expected": "string | null"},
        )
    header_strategy = source.get("headerStrategy", "detect")
    if not isinstance(header_strategy, str):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": f"{field_name}.headerStrategy", "expected": "string"},
        )
    header_row = source.get("headerRow")
    if header_row is not None and (
        isinstance(header_row, bool) or not isinstance(header_row, int)
    ):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": f"{field_name}.headerRow", "expected": "integer | null"},
        )
    return source


def _require_workbook_source_field(
    payload: dict[str, object],
    field_name: str,
) -> dict[str, str]:
    if field_name not in payload:
        raise SheetOpsError(
            "missing_required_field",
            "Required field is missing or invalid",
            {"field": field_name},
        )
    source = payload[field_name]
    if not isinstance(source, dict):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "workbookSource"},
        )
    path = source.get("path")
    if not isinstance(path, str):
        raise SheetOpsError(
            "missing_required_field",
            "Required field is missing or invalid",
            {"field": f"{field_name}.path"},
        )
    return {"path": path}


def _require_workbook_sheet_source_field(
    payload: dict[str, object],
    field_name: str,
) -> dict[str, str]:
    source = _require_workbook_source_field(payload, field_name)
    value = payload[field_name]
    if not isinstance(value, dict):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "workbookSheetSource"},
        )
    sheet = value.get("sheet")
    if not isinstance(sheet, str):
        raise SheetOpsError(
            "missing_required_field",
            "Required field is missing or invalid",
            {"field": f"{field_name}.sheet"},
        )
    return {"path": source["path"], "sheet": sheet}


def _load_table_source(source: dict[str, object]) -> dict[str, object]:
    from zalos_sheetops.io.read_csv_table import read_csv_table
    from zalos_sheetops.io.read_table import read_table

    path = validate_agent_path(str(source["path"]))
    if path.suffix.lower() == ".csv":
        return read_csv_table(path)
    sheet = source.get("sheet")
    header_strategy = source.get("headerStrategy", "detect")
    header_row = source.get("headerRow")
    table = read_table(path, sheet, header_strategy, header_row)
    headers = table["headers"]
    # diff/reconcile address cells by column name, so rehydrate the
    # positional rows into records.
    table["rows"] = [dict(zip(headers, row)) for row in table["rows"]]
    return table


def _validate_string_list_mapping_field(
    payload: dict[str, object],
    field_name: str,
) -> dict[str, list[str]]:
    value = payload.get(field_name, {})
    if not isinstance(value, dict):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "dict[string, list[string]]"},
        )
    for key, item in value.items():
        if not isinstance(key, str) or not isinstance(item, list):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": field_name, "expected": "dict[string, list[string]]"},
            )
        if any(not isinstance(entry, str) for entry in item):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": field_name, "expected": "dict[string, list[string]]"},
            )
    return value


def _require_string_list_mapping_field(
    payload: dict[str, object],
    field_name: str,
) -> dict[str, list[str]]:
    if field_name not in payload:
        raise SheetOpsError(
            "missing_required_field",
            "Required field is missing or invalid",
            {"field": field_name},
        )
    return _validate_string_list_mapping_field(payload, field_name)


def _emit(payload: dict[str, object], pretty: bool) -> None:
    if pretty:
        print(json.dumps(payload, indent=2))
        return
    print(json.dumps(payload, separators=(",", ":")))


# --- dispatcher -------------------------------------------------------------

# (short help for the root listing, long help for the group page)
_GROUPS: dict[str, tuple[str, str]] = {
    "formulas": (
        "Scan formulas and errors.",
        "Scan formulas and formula error tokens.",
    ),
    "cells": (
        "Read bounded cell ranges.",
        "Read bounded cell ranges from worksheets.",
    ),
    "hyperlinks": (
        "Validate targeted cell hyperlinks.",
        "Validate targeted cell hyperlinks.",
    ),
    "styles": (
        "Read and assert cell styles.",
        "Read or assert cell styles, comments, and formulas.",
    ),
    "normalize": (
        "Preview normalization.",
        "Preview deterministic spreadsheet normalization rules.",
    ),
    "doctor": (
        "Check runtime dependencies.",
        "Check local runtime dependencies.",
    ),
    "display": (
        "Scan display overflow risks.",
        "Scan cells that may display as hashes because columns are too narrow.",
    ),
    "compare": (
        "Compare workbook structure or labeled cells.",
        "Compare workbook structure or labeled cells.",
    ),
}

# name path -> (help, needs --json, handler)
_COMMANDS: dict[tuple[str, ...], tuple[str, bool, object]] = {
    ("schema",): (
        "Print the machine-readable command manifest.",
        False,
        _run_schema,
    ),
    ("inspect",): (
        "Inspect workbook metadata; engine defaults to auto/XML for "
        "full-workbook scans.",
        True,
        _run_inspect,
    ),
    ("read-table",): (
        "Read a worksheet or CSV table as JSON records.",
        True,
        _run_read_table,
    ),
    ("validate",): (
        "Validate workbook sheets, columns, keys, and rules.",
        True,
        _run_validate,
    ),
    ("recalc",): (
        "Recalculate a workbook through LibreOffice.",
        True,
        _run_recalc,
    ),
    ("diff",): (
        "Compare two tabular sources by key and value columns.",
        True,
        _run_diff,
    ),
    ("reconcile",): (
        "Reconcile two tabular sources by keys and amounts.",
        True,
        _run_reconcile,
    ),
    ("formulas", "scan"): (
        "Scan formulas and formula error tokens.",
        True,
        _run_formulas_scan,
    ),
    ("cells", "read"): (
        "Read a bounded A1-style cell range. Use engine=openpyxl only for "
        "targeted ranges.",
        True,
        _run_cells_read,
    ),
    ("cells", "search"): (
        "Search bounded cell content or style metadata with deterministic pagination.",
        True,
        _run_cells_search,
    ),
    ("hyperlinks", "assert"): (
        "Validate targeted cell hyperlinks by target, display, and tooltip.",
        True,
        _run_hyperlinks_assert,
    ),
    ("styles", "assert"): (
        "Assert targeted cell styles, formulas, or comments.",
        True,
        _run_styles_assert,
    ),
    ("styles", "read"): (
        "Read styles for a bounded A1-style cell range.",
        True,
        _run_styles_read,
    ),
    ("normalize", "preview"): (
        "Preview column normalization without writing files.",
        True,
        _run_normalize_preview,
    ),
    ("doctor", "libreoffice"): (
        "Check whether LibreOffice is available.",
        False,
        _run_doctor_libreoffice,
    ),
    ("display", "scan"): (
        "Find cells likely to render as hashes because values exceed column width.",
        True,
        _run_display_scan,
    ),
    ("compare", "structure"): (
        "Compare workbook sheet structure.",
        True,
        _run_compare_structure,
    ),
    ("compare", "labeled-cells"): (
        "Compare labeled cells across two workbook sheets.",
        True,
        _run_compare_labeled_cells,
    ),
}

_ROOT_ORDER: list[str] = [
    "schema",
    "inspect",
    "read-table",
    "validate",
    "recalc",
    "diff",
    "reconcile",
    "formulas",
    "cells",
    "hyperlinks",
    "styles",
    "normalize",
    "doctor",
    "display",
    "compare",
]


def _columns(entries: list[tuple[str, str]]) -> list[str]:
    width = max(len(name) for name, _ in entries)
    return [f"  {name.ljust(width)}  {text}" for name, text in entries]


def _help_text(node: tuple[str, ...]) -> str:
    lines: list[str] = []
    if not node:
        lines.append("Usage: sheetops [OPTIONS] COMMAND [ARGS]...")
        lines.append("")
        lines.append(f"  {APP_HELP}")
        lines.append("")
        lines.append("Options:")
        lines.append("  --help  Show this message and exit.")
        lines.append("")
        lines.append("Commands:")
        entries = []
        for name in _ROOT_ORDER:
            if name in _GROUPS:
                entries.append((name, _GROUPS[name][0]))
            else:
                entries.append((name, str(_COMMANDS[(name,)][0])))
        lines.extend(_columns(entries))
    elif node in _COMMANDS:
        help_text, needs_json, _ = _COMMANDS[node]
        lines.append(f"Usage: sheetops {' '.join(node)} [OPTIONS]")
        lines.append("")
        lines.append(f"  {help_text}")
        lines.append("")
        lines.append("Options:")
        entries = []
        if needs_json:
            entries.append(("--json TEXT", f"{JSON_OPTION_HELP} [required]"))
        entries.append(("--pretty", PRETTY_OPTION_HELP))
        entries.append(("--raw-output", RAW_OUTPUT_OPTION_HELP))
        entries.append(("--help", "Show this message and exit."))
        lines.extend(_columns(entries))
    else:
        group = node[0]
        lines.append(f"Usage: sheetops {group} [OPTIONS] COMMAND [ARGS]...")
        lines.append("")
        lines.append(f"  {_GROUPS[group][1]}")
        lines.append("")
        lines.append("Options:")
        lines.append("  --help  Show this message and exit.")
        lines.append("")
        lines.append("Commands:")
        entries = [
            (path[1], str(spec[0]))
            for path, spec in _COMMANDS.items()
            if len(path) == 2 and path[0] == group
        ]
        lines.extend(_columns(entries))
    return "\n".join(lines)


def _usage_error(node: tuple[str, ...], message: str) -> int:
    name = " ".join(("sheetops", *node))
    if node in _COMMANDS:
        usage = f"Usage: {name} [OPTIONS]"
    else:
        usage = f"Usage: {name} [OPTIONS] COMMAND [ARGS]..."
    print(
        f"{usage}\nTry '{name} --help' for help.\n\nError: {message}",
        file=sys.stderr,
    )
    return 2


def _run_leaf(node: tuple[str, ...], tokens: list[str]) -> int:
    _, needs_json, handler = _COMMANDS[node]
    json_payload: str | None = None
    pretty = False
    index = 0
    while index < len(tokens):
        token = tokens[index]
        if token == "--help":
            print(_help_text(node))
            return 0
        if token == "--raw-output":
            # Deprecated alias: compact JSON is the default output.
            index += 1
            continue
        if token == "--pretty":
            pretty = True
            index += 1
            continue
        if needs_json and token == "--json":
            if index + 1 >= len(tokens):
                return _usage_error(node, "Option '--json' requires an argument.")
            json_payload = tokens[index + 1]
            index += 2
            continue
        if needs_json and token.startswith("--json="):
            json_payload = token[len("--json=") :]
            index += 1
            continue
        if token.startswith("-"):
            return _usage_error(node, f"No such option: {token}")
        return _usage_error(node, f"Got unexpected extra argument ({token})")
    if needs_json and json_payload is None:
        return _usage_error(node, "Missing option '--json'.")
    run = handler
    return int(run(json_payload, pretty))


def app(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:]) if argv is None else list(argv)
    if not args:
        print(_help_text(()))
        return 2

    path: list[str] = []
    index = 0
    while index < len(args) and not args[index].startswith("-"):
        word = args[index]
        candidate = (*path, word)
        if candidate in _COMMANDS:
            return _run_leaf(candidate, args[index + 1 :])
        if not path and word in _GROUPS:
            path.append(word)
            index += 1
            continue
        return _usage_error(tuple(path), f"No such command '{word}'.")

    node = tuple(path)
    remaining = args[index:]
    if "--help" in remaining:
        print(_help_text(node))
        return 0
    if path:
        return _usage_error(node, "Missing command.")
    return _usage_error(node, f"No such option: {remaining[0]}")


def main() -> int:
    return app()


if __name__ == "__main__":
    raise SystemExit(app())
