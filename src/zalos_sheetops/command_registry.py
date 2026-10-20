from zalos_sheetops.contracts import error_response as _error_response
from zalos_sheetops.contracts import success_response as _success_response


def command_manifest() -> dict[str, object]:
    return {
        "commands": [
            {
                "name": "check-totals",
                "status": "available",
                "input": {
                    "left": "non-empty list[decimal string]",
                    "right": "non-empty list[decimal string]",
                    "decimalPlaces": "required integer 0..12",
                    "rounding": "half_up (default) | half_even",
                    "sourceTolerance": (
                        "non-negative decimal string (source only; default 0)"
                    ),
                },
            },
            {"name": "schema", "status": "available", "input": {}},
            {
                "name": "inspect",
                "status": "available",
                "input": {
                    "path": "string",
                    "engine": "auto | xml | openpyxl",
                    "maxFormulaCells": "positive integer | null",
                    "maxMergedRanges": "positive integer | null",
                    "maxGraphics": "positive integer | null",
                    "maxDefinedNames": "positive integer | null",
                },
            },
            {
                "name": "read-table",
                "status": "available",
                "input": {
                    "path": "string",
                    "sheet": "string | null",
                    "headerStrategy": "detect | row",
                    "headerRow": "integer | null",
                    "maxRows": "positive integer | null",
                    "valueMode": "raw | typed (XML reads; default raw)",
                    "encoding": "CSV encoding | auto (BOM, otherwise UTF-8)",
                    "delimiter": "CSV single-character delimiter",
                },
            },
            {
                "name": "profile-table",
                "status": "available",
                "input": {
                    "path": "string",
                    "sheet": "string | null (XLSX only)",
                    "headerStrategy": "detect | row",
                    "headerRow": "integer | null",
                    "valueMode": "raw | typed (XML reads; default raw)",
                    "encoding": "CSV encoding | auto (BOM, otherwise UTF-8)",
                    "delimiter": "CSV single-character delimiter",
                    "keyColumns": "list[string] of normalized headers | null",
                    "amountColumns": "list[string] of normalized headers | null",
                    "maxRows": "must be omitted or null; complete population only",
                },
                "output": (
                    "source metadata and profile: counts, column statistics, "
                    "exact totals, checks; no source rows"
                ),
            },
            {
                "name": "cells read",
                "status": "available",
                "input": {
                    "path": "string",
                    "sheet": "string",
                    "range": "A1-style range",
                    "include": (
                        "list[values | formulas | cachedValues | "
                        "numberFormats | dataTypes | sourceValues]"
                    ),
                    "engine": "auto | xml | openpyxl",
                    "valueMode": "raw | typed (XML reads; default raw)",
                },
            },
            {
                "name": "verify-edits",
                "status": "available",
                "input": {
                    "originalPath": "string",
                    "path": "string",
                    "allowedRanges": "dict[string, list[bounded A1 range]]",
                    "structuralSheets": "list[string] (authorized structure edits)",
                    "evaluatedPath": "string | null (caller evaluated copy)",
                },
            },
            {
                "name": "cells search",
                "status": "available",
                "input": {
                    "path": "string",
                    "query": "non-empty string",
                    "sheet": "string | null",
                    "range": "A1-style range | null",
                    "fields": (
                        "non-empty list[values | formulas | cachedValues | "
                        "numberFormats | fontColors | fillColors]"
                    ),
                    "regex": "boolean",
                    "caseSensitive": "boolean",
                    "includeHidden": "boolean",
                    "limit": "positive integer | null",
                    "offset": "non-negative integer | null",
                    "maxCells": "positive integer | null",
                },
            },
            {
                "name": "hyperlinks assert",
                "status": "available",
                "input": {
                    "path": "string",
                    "sheet": "string",
                    "checks": "non-empty list[hyperlinkCheck]",
                },
            },
            {
                "name": "styles assert",
                "status": "available",
                "input": {
                    "path": "string",
                    "sheet": "string",
                    "checks": "non-empty list[styleCheck]",
                },
            },
            {
                "name": "styles read",
                "status": "available",
                "input": {
                    "path": "string",
                    "sheet": "string",
                    "range": "A1-style range",
                    "maxCells": "positive integer | null",
                },
            },
            {
                "name": "validate",
                "status": "available",
                "input": {
                    "path": "string",
                    "requiredSheets": "list[string]",
                    "requiredColumns": "dict[string, list[string]]",
                    "uniqueKeys": "dict[string, list[string]]",
                    "sheets": "list[sheetValidationRule] | null",
                },
            },
            {
                "name": "formulas scan",
                "status": "available",
                "input": {
                    "path": "string",
                    "sheet": "string | null",
                    "range": "A1-style range | null",
                    "maxResults": "positive integer | null",
                    "engine": "auto | xml | openpyxl",
                    "include": "list[formulas | errors] | null",
                    "failFast": "boolean",
                },
            },
            {
                "name": "display scan",
                "status": "available",
                "input": {
                    "path": "string",
                    "sheet": "string | null",
                    "range": "A1-style range | null",
                    "maxResults": "positive integer | null",
                    "maxCells": "positive integer | null",
                    "includeHidden": "boolean",
                },
            },
            {
                "name": "compare structure",
                "status": "available",
                "input": {
                    "expected": "workbookSource",
                    "actual": "workbookSource",
                    "compareFormulaCounts": "boolean",
                },
            },
            {
                "name": "compare labeled-cells",
                "status": "available",
                "input": {
                    "expected": "workbookSheetSource",
                    "actual": "workbookSheetSource",
                    "labelColumn": "column letter",
                    "valueColumn": "column letter",
                    "labels": "non-empty list[string]",
                    "compareFormulas": "boolean",
                    "amountTolerance": "decimal string",
                },
            },
            {
                "name": "normalize preview",
                "status": "available",
                "input": {
                    "path": "string",
                    "sheet": "string | null",
                    "headerStrategy": "detect | row",
                    "headerRow": "integer | null",
                    "columns": "list[columnRule]",
                    "maxRows": "integer | null",
                },
            },
            {"name": "doctor libreoffice", "status": "available", "input": {}},
            {
                "name": "recalc",
                "status": "available",
                "input": {
                    "path": "string",
                    "outputDir": "string",
                    "filename": "string | null",
                    "executable": "string | null",
                },
            },
            {
                "name": "diff",
                "status": "available",
                "input": {
                    "left": "tableSource",
                    "right": "tableSource",
                    "keyColumns": "list[string]",
                    "compareColumns": "list[string] | null",
                    "amountTolerance": "decimal string",
                },
            },
            {
                "name": "reconcile",
                "status": "available",
                "input": {
                    "left": "tableSource",
                    "right": "tableSource",
                    "leftKeyColumns": "list[string]",
                    "rightKeyColumns": "list[string]",
                    "leftAmountColumn": "string",
                    "rightAmountColumn": "string",
                    "amountTolerance": "decimal string",
                },
            },
        ]
    }


def success_response(
    command: str,
    data: dict[str, object],
    warnings: list[object] | None = None,
    artifacts: list[object] | None = None,
) -> dict[str, object]:
    return _success_response(command, data, warnings=warnings, artifacts=artifacts)


def error_response(
    command: str,
    error_code: str,
    message: str,
    details: dict[str, object] | None = None,
    warnings: list[object] | None = None,
    artifacts: list[object] | None = None,
) -> dict[str, object]:
    return _error_response(
        command,
        error_code,
        message,
        details=details,
        warnings=warnings,
        artifacts=artifacts,
    )
