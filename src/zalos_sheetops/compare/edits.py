"""Check edit scope and results in a separately evaluated formula copy.

This is a verifier, not a formula evaluator. The caller must evaluate the copy
in Excel or LibreOffice; a saved cache alone does not establish freshness.
"""

from copy import copy
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.etree.ElementTree import tostring
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.xlsx_xml import MAIN_NS, XlsxWorkbook

# Excel storage names documented at https://xlsxwriter.readthedocs.io/working_with_formulas.html
FUNCTION_STORAGE_NAMES = {
    "ACOT": "_xlfn.ACOT",
    "ACOTH": "_xlfn.ACOTH",
    "AGGREGATE": "_xlfn.AGGREGATE",
    "ANCHORARRAY": "_xlfn.ANCHORARRAY",
    "ARABIC": "_xlfn.ARABIC",
    "ARRAYTOTEXT": "_xlfn.ARRAYTOTEXT",
    "BASE": "_xlfn.BASE",
    "BETA.DIST": "_xlfn.BETA.DIST",
    "BETA.INV": "_xlfn.BETA.INV",
    "BINOM.DIST": "_xlfn.BINOM.DIST",
    "BINOM.DIST.RANGE": "_xlfn.BINOM.DIST.RANGE",
    "BINOM.INV": "_xlfn.BINOM.INV",
    "BITAND": "_xlfn.BITAND",
    "BITLSHIFT": "_xlfn.BITLSHIFT",
    "BITOR": "_xlfn.BITOR",
    "BITRSHIFT": "_xlfn.BITRSHIFT",
    "BITXOR": "_xlfn.BITXOR",
    "BYCOL": "_xlfn.BYCOL",
    "BYROW": "_xlfn.BYROW",
    "CEILING.MATH": "_xlfn.CEILING.MATH",
    "CEILING.PRECISE": "_xlfn.CEILING.PRECISE",
    "CHISQ.DIST": "_xlfn.CHISQ.DIST",
    "CHISQ.DIST.RT": "_xlfn.CHISQ.DIST.RT",
    "CHISQ.INV": "_xlfn.CHISQ.INV",
    "CHISQ.INV.RT": "_xlfn.CHISQ.INV.RT",
    "CHISQ.TEST": "_xlfn.CHISQ.TEST",
    "CHOOSECOLS": "_xlfn.CHOOSECOLS",
    "CHOOSEROWS": "_xlfn.CHOOSEROWS",
    "COMBINA": "_xlfn.COMBINA",
    "CONCAT": "_xlfn.CONCAT",
    "CONFIDENCE.NORM": "_xlfn.CONFIDENCE.NORM",
    "CONFIDENCE.T": "_xlfn.CONFIDENCE.T",
    "COT": "_xlfn.COT",
    "COTH": "_xlfn.COTH",
    "COVARIANCE.P": "_xlfn.COVARIANCE.P",
    "COVARIANCE.S": "_xlfn.COVARIANCE.S",
    "CSC": "_xlfn.CSC",
    "CSCH": "_xlfn.CSCH",
    "DAYS": "_xlfn.DAYS",
    "DECIMAL": "_xlfn.DECIMAL",
    "DROP": "_xlfn.DROP",
    "ERF.PRECISE": "_xlfn.ERF.PRECISE",
    "ERFC.PRECISE": "_xlfn.ERFC.PRECISE",
    "EXPAND": "_xlfn.EXPAND",
    "EXPON.DIST": "_xlfn.EXPON.DIST",
    "F.DIST": "_xlfn.F.DIST",
    "F.DIST.RT": "_xlfn.F.DIST.RT",
    "F.INV": "_xlfn.F.INV",
    "F.INV.RT": "_xlfn.F.INV.RT",
    "F.TEST": "_xlfn.F.TEST",
    "FILTERXML": "_xlfn.FILTERXML",
    "FLOOR.MATH": "_xlfn.FLOOR.MATH",
    "FLOOR.PRECISE": "_xlfn.FLOOR.PRECISE",
    "FORECAST.ETS": "_xlfn.FORECAST.ETS",
    "FORECAST.ETS.CONFINT": "_xlfn.FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY": "_xlfn.FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT": "_xlfn.FORECAST.ETS.STAT",
    "FORECAST.LINEAR": "_xlfn.FORECAST.LINEAR",
    "FORMULATEXT": "_xlfn.FORMULATEXT",
    "GAMMA": "_xlfn.GAMMA",
    "GAMMA.DIST": "_xlfn.GAMMA.DIST",
    "GAMMA.INV": "_xlfn.GAMMA.INV",
    "GAMMALN.PRECISE": "_xlfn.GAMMALN.PRECISE",
    "GAUSS": "_xlfn.GAUSS",
    "HSTACK": "_xlfn.HSTACK",
    "HYPGEOM.DIST": "_xlfn.HYPGEOM.DIST",
    "IFNA": "_xlfn.IFNA",
    "IFS": "_xlfn.IFS",
    "IMAGE": "_xlfn.IMAGE",
    "IMCOSH": "_xlfn.IMCOSH",
    "IMCOT": "_xlfn.IMCOT",
    "IMCSC": "_xlfn.IMCSC",
    "IMCSCH": "_xlfn.IMCSCH",
    "IMSEC": "_xlfn.IMSEC",
    "IMSECH": "_xlfn.IMSECH",
    "IMSINH": "_xlfn.IMSINH",
    "IMTAN": "_xlfn.IMTAN",
    "ISFORMULA": "_xlfn.ISFORMULA",
    "ISOMITTED": "_xlfn.ISOMITTED",
    "ISOWEEKNUM": "_xlfn.ISOWEEKNUM",
    "LAMBDA": "_xlfn.LAMBDA",
    "LET": "_xlfn.LET",
    "LOGNORM.DIST": "_xlfn.LOGNORM.DIST",
    "LOGNORM.INV": "_xlfn.LOGNORM.INV",
    "MAKEARRAY": "_xlfn.MAKEARRAY",
    "MAP": "_xlfn.MAP",
    "MAXIFS": "_xlfn.MAXIFS",
    "MINIFS": "_xlfn.MINIFS",
    "MODE.MULT": "_xlfn.MODE.MULT",
    "MODE.SNGL": "_xlfn.MODE.SNGL",
    "MUNIT": "_xlfn.MUNIT",
    "NEGBINOM.DIST": "_xlfn.NEGBINOM.DIST",
    "NORM.DIST": "_xlfn.NORM.DIST",
    "NORM.INV": "_xlfn.NORM.INV",
    "NORM.S.DIST": "_xlfn.NORM.S.DIST",
    "NORM.S.INV": "_xlfn.NORM.S.INV",
    "NUMBERVALUE": "_xlfn.NUMBERVALUE",
    "PDURATION": "_xlfn.PDURATION",
    "PERCENTILE.EXC": "_xlfn.PERCENTILE.EXC",
    "PERCENTILE.INC": "_xlfn.PERCENTILE.INC",
    "PERCENTRANK.EXC": "_xlfn.PERCENTRANK.EXC",
    "PERCENTRANK.INC": "_xlfn.PERCENTRANK.INC",
    "PERMUTATIONA": "_xlfn.PERMUTATIONA",
    "PHI": "_xlfn.PHI",
    "POISSON.DIST": "_xlfn.POISSON.DIST",
    "QUARTILE.EXC": "_xlfn.QUARTILE.EXC",
    "QUARTILE.INC": "_xlfn.QUARTILE.INC",
    "QUERYSTRING": "_xlfn.QUERYSTRING",
    "RANDARRAY": "_xlfn.RANDARRAY",
    "RANK.AVG": "_xlfn.RANK.AVG",
    "RANK.EQ": "_xlfn.RANK.EQ",
    "REDUCE": "_xlfn.REDUCE",
    "RRI": "_xlfn.RRI",
    "SCAN": "_xlfn.SCAN",
    "SEC": "_xlfn.SEC",
    "SECH": "_xlfn.SECH",
    "SEQUENCE": "_xlfn.SEQUENCE",
    "SHEET": "_xlfn.SHEET",
    "SHEETS": "_xlfn.SHEETS",
    "SINGLE": "_xlfn.SINGLE",
    "SKEW.P": "_xlfn.SKEW.P",
    "SORTBY": "_xlfn.SORTBY",
    "STDEV.P": "_xlfn.STDEV.P",
    "STDEV.S": "_xlfn.STDEV.S",
    "SWITCH": "_xlfn.SWITCH",
    "T.DIST": "_xlfn.T.DIST",
    "T.DIST.2T": "_xlfn.T.DIST.2T",
    "T.DIST.RT": "_xlfn.T.DIST.RT",
    "T.INV": "_xlfn.T.INV",
    "T.INV.2T": "_xlfn.T.INV.2T",
    "T.TEST": "_xlfn.T.TEST",
    "TAKE": "_xlfn.TAKE",
    "TEXTAFTER": "_xlfn.TEXTAFTER",
    "TEXTBEFORE": "_xlfn.TEXTBEFORE",
    "TEXTJOIN": "_xlfn.TEXTJOIN",
    "TEXTSPLIT": "_xlfn.TEXTSPLIT",
    "TOCOL": "_xlfn.TOCOL",
    "TOROW": "_xlfn.TOROW",
    "UNICHAR": "_xlfn.UNICHAR",
    "UNICODE": "_xlfn.UNICODE",
    "UNIQUE": "_xlfn.UNIQUE",
    "VALUETOTEXT": "_xlfn.VALUETOTEXT",
    "VAR.P": "_xlfn.VAR.P",
    "VAR.S": "_xlfn.VAR.S",
    "VSTACK": "_xlfn.VSTACK",
    "WEBSERVICE": "_xlfn.WEBSERVICE",
    "WEIBULL.DIST": "_xlfn.WEIBULL.DIST",
    "WRAPCOLS": "_xlfn.WRAPCOLS",
    "WRAPROWS": "_xlfn.WRAPROWS",
    "XLOOKUP": "_xlfn.XLOOKUP",
    "XMATCH": "_xlfn.XMATCH",
    "XOR": "_xlfn.XOR",
    "Z.TEST": "_xlfn.Z.TEST",
    "FILTER": "_xlfn._xlws.FILTER",
    "SORT": "_xlfn._xlws.SORT",
}
CELL_STYLE_FIELDS = (
    "font",
    "fill",
    "border",
    "alignment",
    "protection",
    "number_format",
)
DYNAMIC_FUNCTIONS = {
    "FILTER",
    "UNIQUE",
    "SORT",
    "SORTBY",
    "SEQUENCE",
    "TEXTSPLIT",
    "HSTACK",
    "VSTACK",
    "TAKE",
    "DROP",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "BYROW",
    "BYCOL",
    "MAP",
    "REDUCE",
    "SCAN",
    "MAKEARRAY",
    "RANDARRAY",
    "TOCOL",
    "TOROW",
    "WRAPROWS",
    "WRAPCOLS",
    "EXPAND",
}


def parse_ranges(ranges: object) -> dict[str, list[tuple[int, int, int, int]]]:
    if not isinstance(ranges, dict):
        raise SheetOpsError(
            "invalid_field",
            "allowedRanges must map sheets to bounded ranges",
            {"field": "allowedRanges"},
        )
    parsed = {}
    for sheet, areas in ranges.items():
        if not isinstance(sheet, str) or not isinstance(areas, list):
            raise SheetOpsError(
                "invalid_field", "Invalid allowedRanges", {"field": "allowedRanges"}
            )
        bounds = []
        for area in areas:
            try:
                box = range_boundaries(area)
                if None in box or box[0] > box[2] or box[1] > box[3] or min(box) < 1:
                    raise ValueError(area)
            except (TypeError, ValueError) as error:
                raise SheetOpsError(
                    "invalid_range",
                    "Edit scope must use bounded A1 ranges",
                    {"sheet": sheet, "range": area},
                ) from error
            bounds.append(box)
        parsed[sheet] = bounds
    return parsed


def in_scope(sheet: str, cell: object, scopes: dict) -> bool:
    return any(
        left <= cell.column <= right and top <= cell.row <= bottom
        for left, top, right, bottom in scopes.get(sheet, [])
    )


def formula_text(cell: object) -> str | None:
    if cell is None or cell.data_type != "f":
        return None
    if isinstance(cell.value, str):
        return cell.value
    return cell.value.text if isinstance(cell.value, ArrayFormula) else None


def formula_signature(cell: object) -> tuple | None:
    if cell is None or cell.data_type != "f":
        return None
    if isinstance(cell.value, str):
        return ("normal", cell.value)
    return (
        type(cell.value).__name__,
        tuple(sorted(dict(cell.value).items())),
        formula_text(cell),
    )


def dimension_signature(dimension: object) -> tuple:
    # Style IDs are local to each workbook. Resolve the actual components.
    attributes = tuple(
        sorted((key, value) for key, value in dict(dimension).items() if key != "s")
    )
    return (
        attributes,
        *(copy(getattr(dimension, field)) for field in CELL_STYLE_FIELDS),
    )


def range_in_scope(sheet: str, reference: str, scopes: dict) -> bool:
    remaining = [parse_ranges({sheet: [reference]})[sheet][0]]
    for left, top, right, bottom in scopes.get(sheet, []):
        next_remaining = []
        for x1, y1, x2, y2 in remaining:
            a, b, c, d = max(x1, left), max(y1, top), min(x2, right), min(y2, bottom)
            if a > c or b > d:
                next_remaining.append((x1, y1, x2, y2))
                continue
            if x1 < a:
                next_remaining.append((x1, y1, a - 1, y2))
            if c < x2:
                next_remaining.append((c + 1, y1, x2, y2))
            if y1 < b:
                next_remaining.append((a, y1, c, b - 1))
            if d < y2:
                next_remaining.append((a, d + 1, c, y2))
        remaining = next_remaining
    return not remaining


def cell_signature(cell: object) -> tuple:
    if cell is None:
        return (None,)
    value = formula_signature(cell) if cell.data_type == "f" else cell.value
    if value is None and not cell.has_style and not cell.comment and not cell.hyperlink:
        return (None,)
    note = (cell.comment.text, cell.comment.author) if cell.comment else None
    link = (
        (
            cell.hyperlink.target,
            cell.hyperlink.location,
            cell.hyperlink.tooltip,
            cell.hyperlink.display,
        )
        if cell.hyperlink
        else None
    )
    return (
        value,
        cell.data_type,
        *(copy(getattr(cell, field)) for field in CELL_STYLE_FIELDS),
        note,
        link,
    )


def sheet_structure(sheet: object) -> tuple:
    return (
        sheet.sheet_state,
        tuple(str(area) for area in sheet.merged_cells.ranges),
        str(sheet.freeze_panes),
        xml_signature(sheet.auto_filter),
        xml_signature(sheet.data_validations),
        xml_signature(sheet.sheet_properties),
        xml_signature(sheet.sheet_format),
        xml_signature(sheet.protection),
        xml_signature(sheet.print_options),
        xml_signature(sheet.page_margins),
        xml_signature(sheet.page_setup),
        tuple(
            (
                str(key.sqref),
                tuple(
                    (
                        xml_signature(rule),
                        xml_signature(rule.dxf) if rule.dxf is not None else None,
                    )
                    for rule in sheet.conditional_formatting[key]
                ),
            )
            for key in sheet.conditional_formatting
        ),
        str(sheet.print_area),
        sheet.print_title_rows,
        sheet.print_title_cols,
        tuple(
            sorted(
                (key, xml_signature(value))
                for key, value in sheet.defined_names.items()
            )
        ),
        tuple(
            (key, dimension_signature(value))
            for key, value in sheet.row_dimensions.items()
        ),
        tuple(
            (key, dimension_signature(value))
            for key, value in sheet.column_dimensions.items()
        ),
        tuple((table.name, xml_signature(table)) for table in sheet.tables.values()),
    )


def xml_signature(value: object) -> bytes:
    return tostring(value.to_tree())


def normalized_formula(text: str | None) -> str:
    # Prefixes describe function storage, not string literals or sheet names.
    tokens = Tokenizer(text or "=").items
    parts = []
    index = 0
    while index < len(tokens):
        token = tokens[index]
        value = token.value
        if token.type == "FUNC":
            value = value.replace("_xlfn.", "").replace("_xlws.", "").upper()
            # LibreOffice serializes Excel's boolean constants as zero-arg
            # functions. Keep quoted text and reference operands untouched.
            if (
                value in {"TRUE(", "FALSE("}
                and index + 1 < len(tokens)
                and tokens[index + 1].subtype == "CLOSE"
            ):
                parts.append(value[:-1])
                index += 2
                continue
        parts.append(value)
        index += 1
    return "".join(parts)


def protected_parts(path: Path) -> dict[str, bytes]:
    with ZipFile(path) as archive:
        return {
            name: archive.read(name)
            for name in archive.namelist()
            if name.startswith(
                (
                    "xl/media/",
                    "xl/charts/",
                    "xl/drawings/",
                    "xl/embeddings/",
                    "xl/activeX/",
                    "xl/theme/",
                    "customXml/",
                )
            )
            or "vbaProject" in name
        }


def empty_string_results(path: Path) -> set[tuple[str, str]]:
    results = set()
    with XlsxWorkbook(path) as book:
        for sheet in book.worksheets:
            with book.archive.open(sheet.path) as stream:
                for _, cell in ET.iterparse(stream, events=("end",)):
                    if cell.tag != f"{{{MAIN_NS}}}c":
                        continue
                    value = cell.find(f"{{{MAIN_NS}}}v")
                    if (
                        cell.get("t") == "str"
                        and value is not None
                        and value.text in (None, "")
                    ):
                        results.add((sheet.name, cell.get("r")))
                    cell.clear()
    return results


def formula_issues(sheet: str, cell: object, evaluated: object) -> list[dict]:
    formula = formula_text(cell)
    reasons = []
    try:
        functions = [
            token.value[:-1]
            for token in Tokenizer(formula).items
            if token.type == "FUNC" and token.subtype == "OPEN"
        ]
    except TokenizerError as error:
        return [
            {
                "sheet": sheet,
                "cell": cell.coordinate,
                "reason": "invalid_formula_syntax",
                "message": str(error),
            }
        ]
    if any(
        name.upper().removeprefix("_XLFN.").removeprefix("_XLWS.")
        in FUNCTION_STORAGE_NAMES
        and name.upper()
        != FUNCTION_STORAGE_NAMES[
            name.upper().removeprefix("_XLFN.").removeprefix("_XLWS.")
        ].upper()
        for name in functions
    ):
        reasons.append("missing_function_prefix")
    if any(
        name.upper().split(".")[-1] in DYNAMIC_FUNCTIONS for name in functions
    ) and not isinstance(cell.value, ArrayFormula):
        reasons.append("array_storage_required")
    if evaluated is None:
        reasons.append("evaluation_required")
    elif sheet not in evaluated.sheetnames:
        reasons.append("evaluation_sheet_missing")
    else:
        evaluated_cell = evaluated[sheet]._cells.get((cell.row, cell.column))
        other_formula = formula_text(evaluated_cell)
        if (
            type(cell.value) is not type(evaluated_cell.value)
            if evaluated_cell is not None
            else True
        ) or (
            isinstance(cell.value, ArrayFormula)
            and cell.value.ref != evaluated_cell.value.ref
        ):
            reasons.append("formula_storage_mismatch")
        if normalized_formula(formula) != normalized_formula(other_formula):
            reasons.append("formula_mismatch")
    return [
        {"sheet": sheet, "cell": cell.coordinate, "reason": reason}
        for reason in reasons
    ]


def cached_formula_issues(
    sheet: str, cell: object, cached: object, blank_results: set, errors: list
) -> list[dict]:
    anchor = cached[sheet][cell.coordinate]
    issues = []
    if anchor.data_type == "e":
        issues.append(
            {
                "sheet": sheet,
                "cell": cell.coordinate,
                "reason": "formula_error",
                "value": anchor.value,
            }
        )
    elif anchor.value is None and (sheet, cell.coordinate) not in blank_results:
        issues.append(
            {"sheet": sheet, "cell": cell.coordinate, "reason": "missing_cached_result"}
        )
    if isinstance(cell.value, ArrayFormula):
        left, top, right, bottom = range_boundaries(cell.value.ref)
        present = sum(
            1
            for result in cached[sheet]._cells.values()
            if left <= result.column <= right
            and top <= result.row <= bottom
            and (
                result.value is not None or (sheet, result.coordinate) in blank_results
            )
        )
        missing = (right - left + 1) * (bottom - top + 1) - present
        if missing:
            issues.append(
                {
                    "sheet": sheet,
                    "cell": cell.coordinate,
                    "reason": "missing_cached_result",
                    "range": cell.value.ref,
                    "missingCount": missing,
                }
            )
        for error in errors:
            if (
                error.coordinate != cell.coordinate
                and left <= error.column <= right
                and top <= error.row <= bottom
            ):
                issues.append(
                    {
                        "sheet": sheet,
                        "cell": error.coordinate,
                        "reason": "formula_error",
                        "value": error.value,
                    }
                )
    return issues


def verify_edits(
    original: Path,
    actual: Path,
    allowed_ranges: object,
    *,
    evaluated_path: Path | None = None,
    structural_sheets: list[str] | None = None,
) -> dict:
    scopes = parse_ranges(allowed_ranges)
    structural = set(structural_sheets or [])
    if not structural <= scopes.keys():
        raise SheetOpsError(
            "invalid_field",
            "Structural sheets must have an explicit edit scope",
            {"field": "structuralSheets"},
        )
    opened = []
    try:
        for path in [original, actual, *([evaluated_path] if evaluated_path else [])]:
            opened.append(
                load_workbook(
                    path, data_only=False, keep_vba=path.suffix.lower() == ".xlsm"
                )
            )
        before, after = opened[:2]
        if not scopes.keys() <= set(before.sheetnames):
            raise SheetOpsError(
                "invalid_sheet",
                "Edit scope contains an unknown original sheet",
                {"sheets": sorted(scopes.keys() - set(before.sheetnames))},
            )
        evaluated = opened[2] if evaluated_path else None
        cached = None
        blank_results = set()
        if evaluated_path:
            cached = load_workbook(evaluated_path, data_only=True)
            opened.append(cached)
            blank_results = empty_string_results(evaluated_path)
        cached_errors = (
            {
                sheet.title: [
                    cell for cell in sheet._cells.values() if cell.data_type == "e"
                ]
                for sheet in cached.worksheets
            }
            if cached is not None
            else {}
        )
        unrelated, structure, issues = [], [], []
        if before.sheetnames != after.sheetnames:
            structure.append("sheet_names_or_order")
        if before.epoch != after.epoch:
            structure.append("date_system")
        if str(before.defined_names) != str(after.defined_names):
            structure.append("defined_names")
        if protected_parts(original) != protected_parts(actual):
            structure.append("embedded_or_drawing_parts")
        for name in dict.fromkeys([*before.sheetnames, *after.sheetnames]):
            if name not in before.sheetnames or name not in after.sheetnames:
                continue
            old, new = before[name], after[name]
            if name not in structural and sheet_structure(old) != sheet_structure(new):
                structure.append(f"sheet_structure:{name}")
            for address in sorted(old._cells.keys() | new._cells.keys()):
                initial, current = old._cells.get(address), new._cells.get(address)
                representative = current if current is not None else initial
                if not in_scope(name, representative, scopes) and cell_signature(
                    initial
                ) != cell_signature(current):
                    unrelated.append({"sheet": name, "cell": representative.coordinate})
                if formula_signature(current) == formula_signature(initial):
                    continue
                cell_issues = []
                if current is not None and current.data_type == "f":
                    if isinstance(current.value, DataTableFormula):
                        cell_issues = [
                            {
                                "sheet": name,
                                "cell": current.coordinate,
                                "reason": "unsupported_data_table_validation",
                            }
                        ]
                    else:
                        cell_issues = formula_issues(name, current, evaluated)
                for target in (initial, current):
                    if (
                        target is not None
                        and isinstance(target.value, (ArrayFormula, DataTableFormula))
                        and not range_in_scope(name, target.value.ref, scopes)
                    ):
                        cell_issues.append(
                            {
                                "sheet": name,
                                "cell": representative.coordinate,
                                "reason": "formula_range_outside_scope",
                                "range": target.value.ref,
                            }
                        )
                if (
                    cached is not None
                    and not cell_issues
                    and current is not None
                    and current.data_type == "f"
                ):
                    cell_issues.extend(
                        cached_formula_issues(
                            name,
                            current,
                            cached,
                            blank_results,
                            cached_errors.get(name, []),
                        )
                    )
                issues.extend(cell_issues)
        return {
            "verified": not (unrelated or structure or issues),
            "unrelatedChangeCount": len(unrelated),
            "unrelatedChanges": unrelated[:50],
            "structureChanges": structure,
            "formulaIssueCount": len(issues),
            "formulaIssues": issues[:50],
            "formulaValidation": "caller_evaluated_copy"
            if evaluated_path
            else "not_evaluated",
        }
    except (OSError, ValueError, KeyError, BadZipFile, ET.ParseError) as error:
        raise SheetOpsError(
            "invalid_workbook", "Workbook could not be verified", {"error": str(error)}
        ) from error
    finally:
        for book in opened:
            book.close()
