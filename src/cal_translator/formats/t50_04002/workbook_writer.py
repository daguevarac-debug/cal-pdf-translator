from __future__ import annotations

import json
import re
import shutil
from dataclasses import asdict, dataclass
from datetime import datetime
from importlib import resources
from pathlib import Path
from typing import Any

from cal_translator.excel.com_backend import open_excel_workbook
from cal_translator.formats.t50_04002 import FORMAT_ID
from cal_translator.formats.t50_04002.validate_template import validate_template

# Excel COM constants used without importing generated type libraries.
_XL_CENTER = -4108
_XL_LEFT = -4131
_XL_TOP = -4160
_XL_CONTINUOUS = 1
_XL_THIN = 2
_XL_EDGE_BOTTOM = 9

_TRACEABILITY_START_ROW = 69
_TRACEABILITY_COUNT = 5
_TRACEABILITY_INSERT_RANGE = "71:73"
_INFO_PRINT_END_ROW = 80

_RESULTS_INSERT_RANGE = "25:83"
_RESULTS_PRINT_END_ROW = 102
_MEASUREMENT_TITLE_ROW = 8
_CHANNEL_LAYOUT: dict[str, tuple[int, int, int]] = {
    "A": (9, 10, 33),
    "B": (34, 35, 58),
    "C": (59, 60, 83),
}
_RESULT_PAGE_BREAKS = (40, 79)
_NOTE_TARGET_ROWS = (86, 87, 89, 90, 91, 93, 95)
_RESULT_FIELDS: tuple[tuple[str, str], ...] = (
    ("B", "range"),
    ("C", "specified_value"),
    ("D", "average_measured_value"),
    ("E", "bias"),
    ("F", "maximum_permissible_error"),
    ("G", "expanded_uncertainty"),
    ("H", "coverage_factor"),
    ("I", "cmc"),
    ("J", "pass_fail"),
    ("K", "tur"),
    ("L", "tar"),
)


@dataclass(slots=True)
class WriteOperation:
    logical_name: str
    sheet: str
    address: str
    value: Any


def _load_cell_map() -> dict[str, Any]:
    try:
        import yaml
    except ImportError as exc:
        raise RuntimeError(
            "PyYAML is required. Install the project with: python -m pip install -e ."
        ) from exc

    resource = resources.files(__package__).joinpath("cell_map.yaml")
    with resource.open("r", encoding="utf-8") as handle:
        payload = yaml.safe_load(handle)
    if not isinstance(payload, dict):
        raise RuntimeError("Invalid T50-04002 cell map.")
    return payload


def _node(cell_map: dict[str, Any], *path: str) -> dict[str, Any]:
    current: Any = cell_map
    for key in path:
        if not isinstance(current, dict) or key not in current:
            raise RuntimeError(f"Missing cell-map path: {'.'.join(path)}")
        current = current[key]
    if not isinstance(current, dict):
        raise RuntimeError(f"Invalid cell-map node: {'.'.join(path)}")
    return current


def _operation(
    cell_map: dict[str, Any],
    logical_name: str,
    path: tuple[str, ...],
    value: Any,
) -> WriteOperation | None:
    if value is None or value == "":
        return None
    target = _node(cell_map, *path)
    if not target.get("writable", False):
        raise RuntimeError(f"Cell-map target is not writable: {logical_name}")
    return WriteOperation(
        logical_name=logical_name,
        sheet=str(target["sheet"]),
        address=str(target["cell"]),
        value=value,
    )


def _number(value: Any) -> float | None:
    if value is None:
        return None
    match = re.search(r"[-+]?\d+(?:[.,]\d+)?", str(value))
    if not match:
        return None
    return float(match.group(0).replace(",", "."))


def _date(value: Any) -> datetime | None:
    if value is None or value == "":
        return None
    try:
        return datetime.strptime(str(value), "%Y-%m-%d")
    except ValueError as exc:
        raise RuntimeError(f"Invalid ISO date in extracted data: {value}") from exc


def _normalize_formula(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "").replace("$", "")).upper()


def _normalize_print_area(value: Any) -> str:
    text = str(value or "").strip()
    if "!" in text:
        text = text.rsplit("!", 1)[-1]
    return text.replace("$", "").replace("'", "").replace(" ", "").upper()


def build_scalar_write_plan(payload: dict[str, Any]) -> list[WriteOperation]:
    if payload.get("format_id") != FORMAT_ID:
        raise RuntimeError(
            f"Unsupported extracted-data format: {payload.get('format_id')}. Expected {FORMAT_ID}."
        )

    cell_map = _load_cell_map()
    equipment = payload.get("equipment") or {}
    customer = payload.get("customer") or {}
    environment = payload.get("environmental_conditions") or {}

    candidates: list[WriteOperation | None] = [
        _operation(cell_map, "certificate_number", ("certificate", "number_input"), payload.get("certificate_number")),
        _operation(cell_map, "internal_code", ("internal_code",), payload.get("internal_code")),
        _operation(cell_map, "equipment.description", ("equipment", "description"), equipment.get("description")),
        _operation(cell_map, "equipment.manufacturer", ("equipment", "manufacturer"), equipment.get("manufacturer")),
        _operation(cell_map, "equipment.serial_number", ("equipment", "serial_number"), equipment.get("serial_number")),
        _operation(cell_map, "equipment.model", ("equipment", "model"), equipment.get("model")),
        _operation(
            cell_map,
            "equipment.customer_identification",
            ("equipment", "customer_identification"),
            equipment.get("customer_identification"),
        ),
        _operation(cell_map, "customer.name", ("customer", "name"), customer.get("name")),
        _operation(cell_map, "customer.address", ("customer", "address"), customer.get("address")),
        _operation(cell_map, "customer.city", ("customer", "city"), customer.get("city")),
        _operation(cell_map, "customer.order_number", ("customer", "order_number"), customer.get("order_number")),
        _operation(cell_map, "dates.received", ("dates", "received"), _date(payload.get("reception_date"))),
        _operation(cell_map, "dates.calibration", ("dates", "calibration"), _date(payload.get("calibration_date"))),
        _operation(cell_map, "dates.issue", ("dates", "issue"), _date(payload.get("issue_date"))),
        _operation(
            cell_map,
            "environment.maximum_temperature.measured",
            ("environment", "maximum_temperature", "measured"),
            _number(environment.get("maximum_temperature")),
        ),
        _operation(
            cell_map,
            "environment.maximum_temperature.correction",
            ("environment", "maximum_temperature", "correction"),
            0.0,
        ),
        _operation(
            cell_map,
            "environment.minimum_temperature.measured",
            ("environment", "minimum_temperature", "measured"),
            _number(environment.get("minimum_temperature")),
        ),
        _operation(
            cell_map,
            "environment.minimum_temperature.correction",
            ("environment", "minimum_temperature", "correction"),
            0.0,
        ),
        _operation(
            cell_map,
            "environment.maximum_humidity.measured",
            ("environment", "maximum_humidity", "measured"),
            _number(environment.get("maximum_relative_humidity")),
        ),
        _operation(
            cell_map,
            "environment.maximum_humidity.correction",
            ("environment", "maximum_humidity", "correction"),
            0.0,
        ),
        _operation(
            cell_map,
            "environment.minimum_humidity.measured",
            ("environment", "minimum_humidity", "measured"),
            _number(environment.get("minimum_relative_humidity")),
        ),
        _operation(
            cell_map,
            "environment.minimum_humidity.correction",
            ("environment", "minimum_humidity", "correction"),
            0.0,
        ),
    ]
    return [operation for operation in candidates if operation is not None]


def _write_value(cell: Any, value: Any) -> None:
    if isinstance(value, datetime):
        cell.Value = value
    else:
        cell.Value2 = value


def _require_complete_payload(payload: dict[str, Any]) -> None:
    traceability = payload.get("traceability") or []
    results = payload.get("results") or []
    notes = payload.get("notes") or []
    method = str(payload.get("calibration_method") or "").strip()

    if len(traceability) != _TRACEABILITY_COUNT:
        raise RuntimeError(
            f"Workbook generation requires {_TRACEABILITY_COUNT} traceability rows; "
            f"found {len(traceability)}."
        )
    if len(results) != 72:
        raise RuntimeError(f"Workbook generation requires 72 result rows; found {len(results)}.")
    counts = {
        channel: sum(str(row.get("channel", "")).upper() == channel for row in results)
        for channel in _CHANNEL_LAYOUT
    }
    if counts != {"A": 24, "B": 24, "C": 24}:
        raise RuntimeError(f"Expected 24 result rows per channel; found {counts}.")
    if not method:
        raise RuntimeError("Workbook generation requires a calibration_method value.")
    if len(notes) != len(_NOTE_TARGET_ROWS):
        raise RuntimeError(
            f"Workbook generation requires {len(_NOTE_TARGET_ROWS)} notes; found {len(notes)}."
        )


def _write_method(info_sheet: Any, method: str) -> None:
    for address in ("B26:H28", "B29:H31"):
        candidate = info_sheet.Range(address)
        if bool(candidate.MergeCells):
            candidate.UnMerge()
    target = info_sheet.Range("B26:H31")
    target.ClearContents()
    target.Merge()
    target.Value2 = method.strip()
    target.WrapText = True
    target.ShrinkToFit = False
    target.HorizontalAlignment = _XL_LEFT
    target.VerticalAlignment = _XL_TOP
    target.Font.Size = 9


def _write_traceability(info_sheet: Any, entries: list[dict[str, Any]]) -> None:
    info_sheet.Rows(_TRACEABILITY_INSERT_RANGE).Insert()

    for row in range(71, 74):
        info_sheet.Range("B70:H70").Copy(Destination=info_sheet.Range(f"B{row}:H{row}"))
        equipment_range = info_sheet.Range(f"B{row}:C{row}")
        if not bool(equipment_range.MergeCells):
            equipment_range.Merge()
        info_sheet.Rows(row).RowHeight = info_sheet.Rows(70).RowHeight

    target = info_sheet.Range("B69:H73")
    target.ClearContents()
    target.NumberFormat = "@"
    target.WrapText = True
    target.VerticalAlignment = _XL_CENTER

    for offset, entry in enumerate(entries):
        row = _TRACEABILITY_START_ROW + offset
        equipment_range = info_sheet.Range(f"B{row}:C{row}")
        if not bool(equipment_range.MergeCells):
            equipment_range.Merge()
        values = {
            "B": entry.get("equipment"),
            "D": entry.get("type"),
            "E": entry.get("internal_number"),
            "F": entry.get("calibrated_by"),
            "G": entry.get("certificate_number"),
            "H": entry.get("calibration_date"),
        }
        for column, value in values.items():
            info_sheet.Range(f"{column}{row}").Value2 = value or ""

    info_sheet.Rows(73).RowHeight = 30
    info_sheet.PageSetup.PrintArea = f"$B$1:$H${_INFO_PRINT_END_ROW}"


def result_display_rows(payload: dict[str, Any]) -> list[tuple[int, dict[str, Any]]]:
    results = payload.get("results") or []
    positioned: list[tuple[int, dict[str, Any]]] = []
    for channel, (_, start_row, end_row) in _CHANNEL_LAYOUT.items():
        channel_rows = [
            row for row in results if str(row.get("channel", "")).upper() == channel
        ]
        expected = end_row - start_row + 1
        if len(channel_rows) != expected:
            raise RuntimeError(
                f"Channel {channel} requires {expected} rows for the controlled layout; "
                f"found {len(channel_rows)}."
            )
        previous_range: str | None = None
        for sheet_row, result in zip(range(start_row, end_row + 1), channel_rows):
            display = dict(result)
            current_range = str(display.get("range") or "").strip()
            display["range"] = current_range if current_range and current_range != previous_range else ""
            if current_range:
                previous_range = current_range
            positioned.append((sheet_row, display))
    return positioned


def _configure_result_title_row(sheet: Any, row: int, text: str) -> None:
    target = sheet.Range(f"B{row}:L{row}")
    if bool(target.MergeCells):
        target.UnMerge()
    target.Merge()
    target.Value2 = text
    target.HorizontalAlignment = _XL_CENTER
    target.VerticalAlignment = _XL_CENTER
    target.Font.Bold = True
    target.WrapText = False
    sheet.Rows(row).RowHeight = 15


def _write_results(result_sheet: Any, payload: dict[str, Any]) -> None:
    # The original sheet has 17 data rows. Insert 59 rows so the controlled
    # result area can hold one measurement title, three channel headings and
    # 72 result rows while leaving the existing notes block intact below it.
    result_sheet.Rows(_RESULTS_INSERT_RANGE).Insert()
    result_sheet.Range("B9:L9").Copy(Destination=result_sheet.Range("B8:L83"))
    result_sheet.Range("B8:L83").ClearContents()
    result_sheet.Range("B8:L83").NumberFormat = "@"
    result_sheet.Range("B8:L83").WrapText = False
    result_sheet.Range("B8:L83").HorizontalAlignment = _XL_CENTER
    result_sheet.Range("B8:L83").VerticalAlignment = _XL_CENTER
    result_sheet.Rows("8:83").RowHeight = 14.5

    _configure_result_title_row(result_sheet, _MEASUREMENT_TITLE_ROW, "Medición de Resistencia")
    for channel, (header_row, _, _) in _CHANNEL_LAYOUT.items():
        _configure_result_title_row(result_sheet, header_row, f"CANAL {channel}")

    for sheet_row, result in result_display_rows(payload):
        for column, field in _RESULT_FIELDS:
            value = result.get(field)
            result_sheet.Range(f"{column}{sheet_row}").Value2 = "" if value is None else str(value)
        if result.get("range"):
            result_sheet.Range(f"B{sheet_row}").Font.Bold = True

    for _, _, last_row in _CHANNEL_LAYOUT.values():
        bottom = result_sheet.Range(f"B{last_row}:L{last_row}").Borders(_XL_EDGE_BOTTOM)
        bottom.LineStyle = _XL_CONTINUOUS
        bottom.Weight = _XL_THIN

    notes = payload.get("notes") or []
    for row, note in zip(_NOTE_TARGET_ROWS, notes):
        result_sheet.Range(f"B{row}").Value2 = str(note)

    result_sheet.PageSetup.PrintArea = f"$B$1:$L${_RESULTS_PRINT_END_ROW}"
    result_sheet.PageSetup.PrintTitleRows = "$3:$7"
    result_sheet.ResetAllPageBreaks()
    for row in _RESULT_PAGE_BREAKS:
        result_sheet.HPageBreaks.Add(Before=result_sheet.Rows(row))


def _check(
    report: dict[str, Any],
    *,
    code: str,
    target: str,
    expected: Any,
    actual: Any,
    passed: bool,
) -> None:
    report["checks"].append(
        {
            "code": code,
            "target": target,
            "expected": expected,
            "actual": actual,
            "passed": passed,
        }
    )
    if not passed:
        report["errors"].append(
            {
                "code": code,
                "target": target,
                "expected": expected,
                "actual": actual,
            }
        )


def validate_generated_workbook(output_path: Path, payload: dict[str, Any]) -> dict[str, Any]:
    report: dict[str, Any] = {"valid": False, "errors": [], "checks": []}
    with open_excel_workbook(output_path, read_only=True) as (_, workbook):
        info = workbook.Worksheets("Información")
        results = workbook.Worksheets("Resultados")

        formulas = {
            "Información!B6": "=B5",
            "Información!H37": "=I19",
            "Información!D43": "=L40+L42",
            "Información!D44": "=L41+L43",
            "Información!H43": "=L44+L46",
            "Información!H44": "=L45+L47",
        }
        for location, expected in formulas.items():
            sheet_name, address = location.split("!", 1)
            cell = workbook.Worksheets(sheet_name).Range(address)
            actual = cell.Formula if bool(cell.HasFormula) else None
            _check(
                report,
                code="required_formula",
                target=location,
                expected=expected,
                actual=actual,
                passed=_normalize_formula(actual) == _normalize_formula(expected),
            )

        _check(
            report,
            code="information_print_area",
            target="Información",
            expected=f"B1:H{_INFO_PRINT_END_ROW}",
            actual=info.PageSetup.PrintArea,
            passed=_normalize_print_area(info.PageSetup.PrintArea)
            == f"B1:H{_INFO_PRINT_END_ROW}",
        )
        _check(
            report,
            code="results_print_area",
            target="Resultados",
            expected=f"B1:L{_RESULTS_PRINT_END_ROW}",
            actual=results.PageSetup.PrintArea,
            passed=_normalize_print_area(results.PageSetup.PrintArea)
            == f"B1:L{_RESULTS_PRINT_END_ROW}",
        )

        method_actual = str(info.Range("B26").Value2 or "").strip()
        method_expected = str(payload.get("calibration_method") or "").strip()
        _check(
            report,
            code="calibration_method",
            target="Información!B26:H31",
            expected=method_expected,
            actual=method_actual,
            passed=method_actual == method_expected,
        )

        for offset, entry in enumerate(payload.get("traceability") or []):
            row = _TRACEABILITY_START_ROW + offset
            expected_values = (
                entry.get("equipment") or "",
                entry.get("type") or "",
                entry.get("internal_number") or "",
                entry.get("calibrated_by") or "",
                entry.get("certificate_number") or "",
                entry.get("calibration_date") or "",
            )
            actual_values = (
                str(info.Range(f"B{row}").Value2 or ""),
                str(info.Range(f"D{row}").Value2 or ""),
                str(info.Range(f"E{row}").Value2 or ""),
                str(info.Range(f"F{row}").Value2 or ""),
                str(info.Range(f"G{row}").Value2 or ""),
                str(info.Range(f"H{row}").Value2 or ""),
            )
            _check(
                report,
                code="traceability_row",
                target=f"Información!B{row}:H{row}",
                expected=expected_values,
                actual=actual_values,
                passed=actual_values == expected_values,
            )

        for sheet_row, expected_result in result_display_rows(payload):
            expected_values = tuple(
                "" if expected_result.get(field) is None else str(expected_result.get(field))
                for _, field in _RESULT_FIELDS
            )
            actual_values = tuple(
                str(results.Range(f"{column}{sheet_row}").Value2 or "")
                for column, _ in _RESULT_FIELDS
            )
            _check(
                report,
                code="result_row",
                target=f"Resultados!B{sheet_row}:L{sheet_row}",
                expected=expected_values,
                actual=actual_values,
                passed=actual_values == expected_values,
            )

        for row, note in zip(_NOTE_TARGET_ROWS, payload.get("notes") or []):
            actual = str(results.Range(f"B{row}").Value2 or "")
            _check(
                report,
                code="note",
                target=f"Resultados!B{row}",
                expected=str(note),
                actual=actual,
                passed=actual == str(note),
            )

    report["valid"] = not report["errors"]
    return report


def build_workbook_prototype(
    template_path: Path,
    data_path: Path,
    output_path: Path,
    *,
    format_id: str = FORMAT_ID,
    overwrite: bool = False,
) -> tuple[dict[str, Any], Path]:
    if format_id != FORMAT_ID:
        raise RuntimeError(f"Unsupported workbook format: {format_id}. Expected {FORMAT_ID}.")
    if not data_path.exists():
        raise FileNotFoundError(f"Extracted JSON file not found: {data_path}")

    payload = json.loads(data_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise RuntimeError("Extracted certificate JSON must contain an object.")
    warnings = payload.get("extraction_warnings") or []
    if warnings:
        raise RuntimeError(
            "Workbook generation requires extraction_warnings to be empty. "
            f"Found {len(warnings)} warning(s)."
        )
    _require_complete_payload(payload)

    preflight = validate_template(template_path, format_id)
    if not preflight["valid"]:
        raise RuntimeError(
            f"Template validation failed with {len(preflight['errors'])} error(s)."
        )

    output_path = output_path.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists() and not overwrite:
        raise RuntimeError(f"Output workbook already exists: {output_path}")
    if output_path.suffix.lower() != ".xlsx":
        raise RuntimeError("Workbook output must use the .xlsx extension.")

    shutil.copy2(template_path.resolve(), output_path)
    plan = build_scalar_write_plan(payload)

    try:
        with open_excel_workbook(output_path, read_only=False) as (excel, workbook):
            if bool(workbook.ReadOnly):
                raise RuntimeError("Copied workbook unexpectedly opened as read-only.")
            for operation in plan:
                cell = workbook.Worksheets(operation.sheet).Range(operation.address)
                _write_value(cell, operation.value)

            info_sheet = workbook.Worksheets("Información")
            result_sheet = workbook.Worksheets("Resultados")
            _write_method(info_sheet, str(payload["calibration_method"]))
            _write_traceability(info_sheet, list(payload["traceability"]))
            _write_results(result_sheet, payload)

            excel.CalculateFullRebuild()
            workbook.Save()
    except Exception:
        output_path.unlink(missing_ok=True)
        raise

    postflight = validate_generated_workbook(output_path, payload)
    report: dict[str, Any] = {
        "status": "expanded_prototype",
        "format_id": format_id,
        "source_template": str(template_path.resolve()),
        "source_data": str(data_path.resolve()),
        "output_workbook": str(output_path),
        "scalar_fields_written": len(plan),
        "method_written": True,
        "traceability_rows_written": len(payload.get("traceability") or []),
        "result_rows_written": len(payload.get("results") or []),
        "notes_written": len(payload.get("notes") or []),
        "write_operations": [asdict(operation) for operation in plan],
        "layout": {
            "information_print_area": f"B1:H{_INFO_PRINT_END_ROW}",
            "results_print_area": f"B1:L{_RESULTS_PRINT_END_ROW}",
            "result_page_breaks_before_rows": list(_RESULT_PAGE_BREAKS),
            "channel_layout": {
                channel: {
                    "header_row": header,
                    "data_start_row": start,
                    "data_end_row": end,
                }
                for channel, (header, start, end) in _CHANNEL_LAYOUT.items()
            },
        },
        "skipped_sections": {
            "controlled_translation": {
                "reason": "English controlled-text replacement is a separate validation stage"
            }
        },
        "postflight_valid": bool(postflight["valid"]),
        "postflight_errors": postflight["errors"],
        "postflight_checks": postflight["checks"],
    }

    report_path = output_path.with_name(f"{output_path.stem}_build_report.json")
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    return report, report_path
