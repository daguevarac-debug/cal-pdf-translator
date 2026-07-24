from __future__ import annotations

import json
import shutil
from dataclasses import asdict
from pathlib import Path
from typing import Any

from cal_translator.excel.com_backend import open_excel_workbook, retry_com_call, set_com_property
from cal_translator.formats.t50_04002 import FORMAT_ID
from cal_translator.formats.t50_04002 import workbook_writer as legacy
from cal_translator.formats.t50_04002 import workbook_writer_v2 as v2
from cal_translator.formats.t50_04002.validate_template import validate_template

_TRACE_START = 69
_TRACE_END = 71
_INFO_PRINT_END = 78
_RESULT_TITLE_ROW = 11
_RESULT_ACCURACY_ROW = 12
_RESULT_RANGE_ROW = 13
_RESULT_START = 14
_RESULT_END = 20
_RESULT_PRINT_END = 43
_NOTE_ROWS = (27, 28, 30, 31, 32, 34, 36)
_RESULT_FIELDS = (
    "specified_value",
    "average_measured_value",
    "bias",
    "maximum_permissible_error",
    "expanded_uncertainty",
    "coverage_factor",
)


def _as_text(value: Any) -> str:
    return "" if value is None else str(value)


def _require_payload(payload: dict[str, Any]) -> None:
    if payload.get("format_id") != FORMAT_ID:
        raise RuntimeError(
            f"Unsupported extracted-data format: {payload.get('format_id')}. Expected {FORMAT_ID}."
        )
    if payload.get("measurement_kind") != "temperature":
        raise RuntimeError(
            f"Temperature workbook writer received measurement_kind={payload.get('measurement_kind')!r}."
        )
    if payload.get("extraction_warnings"):
        raise RuntimeError("Workbook generation requires extraction_warnings to be empty.")
    if len(payload.get("traceability") or []) != 3:
        raise RuntimeError("Temperature workbook generation requires 3 traceability rows.")
    if len(payload.get("results") or []) != 7:
        raise RuntimeError("Temperature workbook generation requires 7 result rows.")
    if len(payload.get("notes") or []) != 7:
        raise RuntimeError("Temperature workbook generation requires 7 certificate notes.")
    if not str(payload.get("calibration_method") or "").strip():
        raise RuntimeError("Temperature workbook generation requires a calibration method.")


def _column_matrix(entries: list[dict[str, Any]], field: str) -> tuple[tuple[str], ...]:
    return tuple((_as_text(entry.get(field)),) for entry in entries)


def _write_traceability(info_sheet: Any, entries: list[dict[str, Any]]) -> None:
    insert_row = retry_com_call(lambda: info_sheet.Rows("71:71"))
    retry_com_call(insert_row.Insert)
    source = retry_com_call(lambda: info_sheet.Range("B70:H70"))
    destination = retry_com_call(lambda: info_sheet.Range("B71:H71"))
    retry_com_call(lambda: source.Copy(Destination=destination))

    target = retry_com_call(lambda: info_sheet.Range(f"B{_TRACE_START}:H{_TRACE_END}"))
    retry_com_call(target.ClearContents)
    set_com_property(target, "NumberFormat", "@")
    set_com_property(target, "WrapText", True)
    set_com_property(target, "VerticalAlignment", legacy._XL_CENTER)

    for row in range(_TRACE_START, _TRACE_END + 1):
        equipment = retry_com_call(lambda row=row: info_sheet.Range(f"B{row}:C{row}"))
        if not bool(retry_com_call(lambda equipment=equipment: equipment.MergeCells)):
            retry_com_call(equipment.Merge)
        set_com_property(retry_com_call(lambda row=row: info_sheet.Rows(row)), "RowHeight", 22)

    mappings = (
        ("B", "equipment"),
        ("D", "type"),
        ("E", "internal_number"),
        ("F", "calibrated_by"),
        ("G", "certificate_number"),
        ("H", "calibration_date"),
    )
    for column, field in mappings:
        destination = retry_com_call(
            lambda column=column: info_sheet.Range(f"{column}{_TRACE_START}:{column}{_TRACE_END}")
        )
        set_com_property(destination, "Value2", _column_matrix(entries, field))

    page_setup = retry_com_call(lambda: info_sheet.PageSetup)
    set_com_property(page_setup, "PrintArea", f"$B$1:$H${_INFO_PRINT_END}")


def _configure_merged_title(result_sheet: Any, row: int, text: str) -> None:
    full = retry_com_call(lambda: result_sheet.Range(f"B{row}:L{row}"))
    if bool(retry_com_call(lambda: full.MergeCells)):
        retry_com_call(full.UnMerge)
    target = retry_com_call(lambda: result_sheet.Range(f"B{row}:H{row}"))
    retry_com_call(target.Merge)
    set_com_property(target, "Value2", text)
    set_com_property(target, "HorizontalAlignment", legacy._XL_CENTER)
    set_com_property(target, "VerticalAlignment", legacy._XL_CENTER)
    set_com_property(target, "WrapText", False)
    set_com_property(retry_com_call(lambda: target.Font), "Bold", True)


def temperature_result_matrix(payload: dict[str, Any]) -> tuple[tuple[str, ...], ...]:
    return tuple(
        tuple(_as_text(row.get(field)) for field in _RESULT_FIELDS)
        for row in payload.get("results") or []
    )


def _write_results(result_sheet: Any, payload: dict[str, Any]) -> None:
    body = retry_com_call(lambda: result_sheet.Range("B8:L24"))
    retry_com_call(body.ClearContents)
    set_com_property(body, "NumberFormat", "@")

    metadata = payload.get("result_metadata") or {}
    _configure_merged_title(
        result_sheet,
        _RESULT_TITLE_ROW,
        str(metadata.get("measurement_title") or "Medición de Temperatura"),
    )

    labels = {
        f"B{_RESULT_ACCURACY_ROW}": "Exactitud",
        f"C{_RESULT_ACCURACY_ROW}": _as_text(metadata.get("accuracy")),
        f"B{_RESULT_RANGE_ROW}": "Desde",
        f"C{_RESULT_RANGE_ROW}": _as_text(metadata.get("from")),
        f"D{_RESULT_RANGE_ROW}": "Hasta",
        f"E{_RESULT_RANGE_ROW}": _as_text(metadata.get("to")),
    }
    for address, value in labels.items():
        cell = retry_com_call(lambda address=address: result_sheet.Range(address))
        set_com_property(cell, "Value2", value)
    for address in (f"B{_RESULT_ACCURACY_ROW}", f"B{_RESULT_RANGE_ROW}", f"D{_RESULT_RANGE_ROW}"):
        font = retry_com_call(lambda address=address: result_sheet.Range(address).Font)
        set_com_property(font, "Bold", True)
    for address in (f"C{_RESULT_RANGE_ROW}", f"E{_RESULT_RANGE_ROW}"):
        font = retry_com_call(lambda address=address: result_sheet.Range(address).Font)
        set_com_property(font, "Color", 255)
        set_com_property(font, "Bold", True)

    target = retry_com_call(
        lambda: result_sheet.Range(f"C{_RESULT_START}:H{_RESULT_END}")
    )
    set_com_property(target, "Value2", temperature_result_matrix(payload))
    set_com_property(target, "HorizontalAlignment", legacy._XL_CENTER)
    set_com_property(target, "VerticalAlignment", legacy._XL_CENTER)
    set_com_property(target, "WrapText", False)
    set_com_property(retry_com_call(lambda: result_sheet.Rows(f"{_RESULT_START}:{_RESULT_END}")), "RowHeight", 14.5)

    for row in _NOTE_ROWS:
        retry_com_call(lambda row=row: result_sheet.Range(f"B{row}").ClearContents())
    for row, note in zip(_NOTE_ROWS, payload.get("notes") or []):
        set_com_property(
            retry_com_call(lambda row=row: result_sheet.Range(f"B{row}")),
            "Value2",
            str(note),
        )

    page_setup = retry_com_call(lambda: result_sheet.PageSetup)
    set_com_property(page_setup, "PrintArea", f"$B$1:$H${_RESULT_PRINT_END}")
    set_com_property(page_setup, "PrintTitleRows", "")
    retry_com_call(result_sheet.ResetAllPageBreaks)


def _matrix(value: Any) -> tuple[tuple[Any, ...], ...]:
    if isinstance(value, tuple):
        if value and not isinstance(value[0], tuple):
            return (tuple(value),)
        return tuple(tuple(row) for row in value)
    return ((value,),)


def validate_temperature_workbook(output_path: Path, payload: dict[str, Any]) -> dict[str, Any]:
    errors: list[dict[str, Any]] = []
    checks: list[dict[str, Any]] = []

    def check(code: str, target: str, expected: Any, actual: Any) -> None:
        passed = actual == expected
        item = {"code": code, "target": target, "expected": expected, "actual": actual, "passed": passed}
        checks.append(item)
        if not passed:
            errors.append(item)

    with open_excel_workbook(output_path, read_only=True) as (_, workbook):
        info = retry_com_call(lambda: workbook.Worksheets("Información"))
        results = retry_com_call(lambda: workbook.Worksheets("Resultados"))
        check(
            "information_print_area",
            "Información",
            f"B1:H{_INFO_PRINT_END}",
            legacy._normalize_print_area(retry_com_call(lambda: info.PageSetup.PrintArea)),
        )
        check(
            "results_print_area",
            "Resultados",
            f"B1:H{_RESULT_PRINT_END}",
            legacy._normalize_print_area(retry_com_call(lambda: results.PageSetup.PrintArea)),
        )
        check(
            "calibration_method",
            "Información!B26",
            str(payload.get("calibration_method") or "").strip(),
            _as_text(retry_com_call(lambda: info.Range("B26").Value2)).strip(),
        )

        trace = _matrix(retry_com_call(lambda: info.Range("B69:H71").Value2))
        columns = (0, 2, 3, 4, 5, 6)
        for offset, entry in enumerate(payload.get("traceability") or []):
            expected = tuple(
                _as_text(entry.get(field))
                for field in (
                    "equipment",
                    "type",
                    "internal_number",
                    "calibrated_by",
                    "certificate_number",
                    "calibration_date",
                )
            )
            actual = tuple(_as_text(trace[offset][index]) for index in columns)
            check("traceability_row", f"Información!B{69 + offset}:H{69 + offset}", expected, actual)

        actual_matrix = tuple(
            tuple(_as_text(value) for value in row)
            for row in _matrix(retry_com_call(lambda: results.Range("C14:H20").Value2))
        )
        check("temperature_results", "Resultados!C14:H20", temperature_result_matrix(payload), actual_matrix)
        metadata = payload.get("result_metadata") or {}
        for address, key in (("C12", "accuracy"), ("C13", "from"), ("E13", "to")):
            check(
                "temperature_metadata",
                f"Resultados!{address}",
                _as_text(metadata.get(key)),
                _as_text(retry_com_call(lambda address=address: results.Range(address).Value2)),
            )
        for row, note in zip(_NOTE_ROWS, payload.get("notes") or []):
            check(
                "note",
                f"Resultados!B{row}",
                str(note),
                _as_text(retry_com_call(lambda row=row: results.Range(f"B{row}").Value2)),
            )

    return {"valid": not errors, "errors": errors, "checks": checks}


def build_temperature_workbook(
    template_path: Path,
    data_path: Path,
    output_path: Path,
    *,
    format_id: str = FORMAT_ID,
    overwrite: bool = False,
) -> tuple[dict[str, Any], Path]:
    if format_id != FORMAT_ID:
        raise RuntimeError(f"Unsupported workbook format: {format_id}. Expected {FORMAT_ID}.")
    payload = json.loads(data_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise RuntimeError("Extracted certificate JSON must contain an object.")
    _require_payload(payload)

    preflight = validate_template(template_path, format_id)
    if not preflight["valid"]:
        raise RuntimeError(
            f"Template validation failed with {len(preflight['errors'])} error(s)."
        )

    output_path = output_path.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.suffix.lower() != ".xlsx":
        raise RuntimeError("Workbook output must use the .xlsx extension.")
    if output_path.exists():
        if not overwrite:
            raise RuntimeError(f"Output workbook already exists: {output_path}")
        v2._remove_with_retry(output_path)

    shutil.copy2(template_path.resolve(), output_path)
    plan = legacy.build_scalar_write_plan(payload)
    try:
        with open_excel_workbook(output_path, read_only=False) as (excel, workbook):
            if bool(retry_com_call(lambda: workbook.ReadOnly)):
                raise RuntimeError("Copied workbook unexpectedly opened as read-only.")
            v2._write_scalar_fields(workbook, plan)
            info = retry_com_call(lambda: workbook.Worksheets("Información"))
            results = retry_com_call(lambda: workbook.Worksheets("Resultados"))
            v2._write_method(info, str(payload["calibration_method"]))
            _write_traceability(info, list(payload["traceability"]))
            _write_results(results, payload)
            retry_com_call(excel.CalculateFullRebuild, attempts=20, initial_delay=0.25)
            retry_com_call(workbook.Save, attempts=20, initial_delay=0.25)
    except Exception as exc:
        try:
            v2._remove_with_retry(output_path)
        except RuntimeError as cleanup_exc:
            raise RuntimeError(f"{exc}\nCleanup also failed: {cleanup_exc}") from exc
        raise

    postflight = validate_temperature_workbook(output_path, payload)
    report: dict[str, Any] = {
        "status": "temperature_workbook",
        "format_id": format_id,
        "measurement_kind": "temperature",
        "source_template": str(template_path.resolve()),
        "source_data": str(data_path.resolve()),
        "output_workbook": str(output_path),
        "scalar_fields_written": len(plan),
        "method_written": True,
        "traceability_rows_written": 3,
        "result_rows_written": 7,
        "notes_written": 7,
        "source_issues": list(payload.get("source_issues") or []),
        "write_operations": [asdict(operation) for operation in plan],
        "layout": {
            "information_print_area": f"B1:H{_INFO_PRINT_END}",
            "results_print_area": f"B1:H{_RESULT_PRINT_END}",
            "temperature_result_rows": f"14:20",
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
