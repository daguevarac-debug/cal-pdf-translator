from __future__ import annotations

import json
import shutil
import time
from dataclasses import asdict
from pathlib import Path
from typing import Any

from cal_translator.excel.com_backend import (
    open_excel_workbook,
    retry_com_call,
    set_com_property,
)
from cal_translator.formats.t50_04002 import FORMAT_ID
from cal_translator.formats.t50_04002 import workbook_writer as legacy
from cal_translator.formats.t50_04002.validate_template import validate_template


def _as_text(value: Any) -> str:
    return "" if value is None else str(value)


def _set_range_value(target: Any, value: Any) -> None:
    set_com_property(target, "Value2", value)


def _set_range_property(target: Any, name: str, value: Any) -> None:
    set_com_property(target, name, value)


def _write_scalar_fields(workbook: Any, plan: list[legacy.WriteOperation]) -> None:
    for operation in plan:
        sheet = retry_com_call(lambda name=operation.sheet: workbook.Worksheets(name))
        cell = retry_com_call(lambda s=sheet, address=operation.address: s.Range(address))
        if hasattr(operation.value, "year") and hasattr(operation.value, "month"):
            set_com_property(cell, "Value", operation.value)
        else:
            _set_range_value(cell, operation.value)


def _write_method(info_sheet: Any, method: str) -> None:
    for address in ("B26:H28", "B29:H31"):
        candidate = retry_com_call(lambda address=address: info_sheet.Range(address))
        if bool(retry_com_call(lambda c=candidate: c.MergeCells)):
            retry_com_call(candidate.UnMerge)

    target = retry_com_call(lambda: info_sheet.Range("B26:H31"))
    retry_com_call(target.ClearContents)
    retry_com_call(target.Merge)
    _set_range_value(target, method.strip())
    _set_range_property(target, "WrapText", True)
    _set_range_property(target, "ShrinkToFit", False)
    _set_range_property(target, "HorizontalAlignment", legacy._XL_LEFT)
    _set_range_property(target, "VerticalAlignment", legacy._XL_TOP)
    font = retry_com_call(lambda: target.Font)
    _set_range_property(font, "Size", 9)


def _column_matrix(entries: list[dict[str, Any]], field: str) -> tuple[tuple[str], ...]:
    return tuple((_as_text(entry.get(field)),) for entry in entries)


def _write_traceability(info_sheet: Any, entries: list[dict[str, Any]]) -> None:
    rows = retry_com_call(lambda: info_sheet.Rows(legacy._TRACEABILITY_INSERT_RANGE))
    retry_com_call(rows.Insert)

    source = retry_com_call(lambda: info_sheet.Range("B70:H70"))
    source_height = retry_com_call(lambda: info_sheet.Rows(70).RowHeight)
    for row in range(71, 74):
        destination = retry_com_call(lambda row=row: info_sheet.Range(f"B{row}:H{row}"))
        retry_com_call(lambda destination=destination: source.Copy(Destination=destination))
        equipment_range = retry_com_call(lambda row=row: info_sheet.Range(f"B{row}:C{row}"))
        if not bool(retry_com_call(lambda target=equipment_range: target.MergeCells)):
            retry_com_call(equipment_range.Merge)
        set_com_property(retry_com_call(lambda row=row: info_sheet.Rows(row)), "RowHeight", source_height)

    target = retry_com_call(lambda: info_sheet.Range("B69:H73"))
    retry_com_call(target.ClearContents)
    _set_range_property(target, "NumberFormat", "@")
    _set_range_property(target, "WrapText", True)
    _set_range_property(target, "VerticalAlignment", legacy._XL_CENTER)

    for row in range(legacy._TRACEABILITY_START_ROW, legacy._TRACEABILITY_START_ROW + len(entries)):
        equipment_range = retry_com_call(lambda row=row: info_sheet.Range(f"B{row}:C{row}"))
        if not bool(retry_com_call(lambda target=equipment_range: target.MergeCells)):
            retry_com_call(equipment_range.Merge)

    mappings = (
        ("B", "equipment"),
        ("D", "type"),
        ("E", "internal_number"),
        ("F", "calibrated_by"),
        ("G", "certificate_number"),
        ("H", "calibration_date"),
    )
    end_row = legacy._TRACEABILITY_START_ROW + len(entries) - 1
    for column, field in mappings:
        destination = retry_com_call(
            lambda column=column: info_sheet.Range(
                f"{column}{legacy._TRACEABILITY_START_ROW}:{column}{end_row}"
            )
        )
        _set_range_value(destination, _column_matrix(entries, field))

    set_com_property(retry_com_call(lambda: info_sheet.Rows(73)), "RowHeight", 30)
    page_setup = retry_com_call(lambda: info_sheet.PageSetup)
    _set_range_property(page_setup, "PrintArea", f"$B$1:$H${legacy._INFO_PRINT_END_ROW}")


def _configure_title_row(sheet: Any, row: int, text: str) -> None:
    target = retry_com_call(lambda: sheet.Range(f"B{row}:L{row}"))
    if bool(retry_com_call(lambda: target.MergeCells)):
        retry_com_call(target.UnMerge)
    retry_com_call(target.Merge)
    _set_range_value(target, text)
    _set_range_property(target, "HorizontalAlignment", legacy._XL_CENTER)
    _set_range_property(target, "VerticalAlignment", legacy._XL_CENTER)
    _set_range_property(retry_com_call(lambda: target.Font), "Bold", True)
    _set_range_property(target, "WrapText", False)
    set_com_property(retry_com_call(lambda: sheet.Rows(row)), "RowHeight", 15)


def channel_result_matrix(
    payload: dict[str, Any], channel: str
) -> tuple[tuple[str, ...], ...]:
    positioned = dict(legacy.result_display_rows(payload))
    _, start_row, end_row = legacy._CHANNEL_LAYOUT[channel]
    return tuple(
        tuple(_as_text(positioned[row].get(field)) for _, field in legacy._RESULT_FIELDS)
        for row in range(start_row, end_row + 1)
    )


def _write_results(result_sheet: Any, payload: dict[str, Any]) -> None:
    insert_rows = retry_com_call(lambda: result_sheet.Rows(legacy._RESULTS_INSERT_RANGE))
    retry_com_call(insert_rows.Insert)

    source = retry_com_call(lambda: result_sheet.Range("B9:L9"))
    destination = retry_com_call(lambda: result_sheet.Range("B8:L83"))
    retry_com_call(lambda: source.Copy(Destination=destination))
    retry_com_call(destination.ClearContents)
    _set_range_property(destination, "NumberFormat", "@")
    _set_range_property(destination, "WrapText", False)
    _set_range_property(destination, "HorizontalAlignment", legacy._XL_CENTER)
    _set_range_property(destination, "VerticalAlignment", legacy._XL_CENTER)
    set_com_property(retry_com_call(lambda: result_sheet.Rows("8:83")), "RowHeight", 14.5)

    _configure_title_row(result_sheet, legacy._MEASUREMENT_TITLE_ROW, "Medición de Resistencia")
    for channel, (header_row, start_row, end_row) in legacy._CHANNEL_LAYOUT.items():
        _configure_title_row(result_sheet, header_row, f"CANAL {channel}")
        channel_range = retry_com_call(
            lambda start_row=start_row, end_row=end_row: result_sheet.Range(
                f"B{start_row}:L{end_row}"
            )
        )
        _set_range_value(channel_range, channel_result_matrix(payload, channel))

    for sheet_row, result in legacy.result_display_rows(payload):
        if result.get("range"):
            range_cell = retry_com_call(lambda sheet_row=sheet_row: result_sheet.Range(f"B{sheet_row}"))
            _set_range_property(retry_com_call(lambda: range_cell.Font), "Bold", True)

    for _, _, last_row in legacy._CHANNEL_LAYOUT.values():
        bottom_range = retry_com_call(lambda last_row=last_row: result_sheet.Range(f"B{last_row}:L{last_row}"))
        borders = retry_com_call(lambda: bottom_range.Borders(legacy._XL_EDGE_BOTTOM))
        _set_range_property(borders, "LineStyle", legacy._XL_CONTINUOUS)
        _set_range_property(borders, "Weight", legacy._XL_THIN)

    for row, note in zip(legacy._NOTE_TARGET_ROWS, payload.get("notes") or []):
        note_cell = retry_com_call(lambda row=row: result_sheet.Range(f"B{row}"))
        _set_range_value(note_cell, str(note))

    page_setup = retry_com_call(lambda: result_sheet.PageSetup)
    _set_range_property(page_setup, "PrintArea", f"$B$1:$L${legacy._RESULTS_PRINT_END_ROW}")
    _set_range_property(page_setup, "PrintTitleRows", "$3:$7")
    retry_com_call(result_sheet.ResetAllPageBreaks)
    for row in legacy._RESULT_PAGE_BREAKS:
        breaks = retry_com_call(lambda: result_sheet.HPageBreaks)
        before = retry_com_call(lambda row=row: result_sheet.Rows(row))
        retry_com_call(lambda breaks=breaks, before=before: breaks.Add(Before=before))


def _check(
    report: dict[str, Any],
    *,
    code: str,
    target: str,
    expected: Any,
    actual: Any,
    passed: bool,
) -> None:
    item = {
        "code": code,
        "target": target,
        "expected": expected,
        "actual": actual,
        "passed": passed,
    }
    report["checks"].append(item)
    if not passed:
        report["errors"].append(item)


def _matrix_values(value: Any) -> tuple[tuple[Any, ...], ...]:
    if isinstance(value, tuple):
        if value and not isinstance(value[0], tuple):
            return (tuple(value),)
        return tuple(tuple(row) for row in value)
    return ((value,),)


def validate_generated_workbook(output_path: Path, payload: dict[str, Any]) -> dict[str, Any]:
    report: dict[str, Any] = {"valid": False, "errors": [], "checks": []}
    with open_excel_workbook(output_path, read_only=True) as (_, workbook):
        info = retry_com_call(lambda: workbook.Worksheets("Información"))
        results = retry_com_call(lambda: workbook.Worksheets("Resultados"))

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
            sheet = retry_com_call(lambda sheet_name=sheet_name: workbook.Worksheets(sheet_name))
            cell = retry_com_call(lambda sheet=sheet, address=address: sheet.Range(address))
            has_formula = bool(retry_com_call(lambda cell=cell: cell.HasFormula))
            actual = retry_com_call(lambda cell=cell: cell.Formula) if has_formula else None
            _check(
                report,
                code="required_formula",
                target=location,
                expected=expected,
                actual=actual,
                passed=legacy._normalize_formula(actual) == legacy._normalize_formula(expected),
            )

        info_page_setup = retry_com_call(lambda: info.PageSetup)
        info_area = retry_com_call(lambda: info_page_setup.PrintArea)
        _check(
            report,
            code="information_print_area",
            target="Información",
            expected=f"B1:H{legacy._INFO_PRINT_END_ROW}",
            actual=info_area,
            passed=legacy._normalize_print_area(info_area)
            == f"B1:H{legacy._INFO_PRINT_END_ROW}",
        )

        result_page_setup = retry_com_call(lambda: results.PageSetup)
        result_area = retry_com_call(lambda: result_page_setup.PrintArea)
        _check(
            report,
            code="results_print_area",
            target="Resultados",
            expected=f"B1:L{legacy._RESULTS_PRINT_END_ROW}",
            actual=result_area,
            passed=legacy._normalize_print_area(result_area)
            == f"B1:L{legacy._RESULTS_PRINT_END_ROW}",
        )

        method_cell = retry_com_call(lambda: info.Range("B26"))
        method_actual = _as_text(retry_com_call(lambda: method_cell.Value2)).strip()
        method_expected = _as_text(payload.get("calibration_method")).strip()
        _check(
            report,
            code="calibration_method",
            target="Información!B26:H31",
            expected=method_expected,
            actual=method_actual,
            passed=method_actual == method_expected,
        )

        trace_range = retry_com_call(lambda: info.Range("B69:H73"))
        trace_values = _matrix_values(retry_com_call(lambda: trace_range.Value2))
        trace_columns = (0, 2, 3, 4, 5, 6)
        for offset, entry in enumerate(payload.get("traceability") or []):
            expected_values = (
                _as_text(entry.get("equipment")),
                _as_text(entry.get("type")),
                _as_text(entry.get("internal_number")),
                _as_text(entry.get("calibrated_by")),
                _as_text(entry.get("certificate_number")),
                _as_text(entry.get("calibration_date")),
            )
            actual_values = tuple(_as_text(trace_values[offset][index]) for index in trace_columns)
            _check(
                report,
                code="traceability_row",
                target=f"Información!B{legacy._TRACEABILITY_START_ROW + offset}:H{legacy._TRACEABILITY_START_ROW + offset}",
                expected=expected_values,
                actual=actual_values,
                passed=actual_values == expected_values,
            )

        for channel, (_, start_row, end_row) in legacy._CHANNEL_LAYOUT.items():
            expected_matrix = channel_result_matrix(payload, channel)
            result_range = retry_com_call(
                lambda start_row=start_row, end_row=end_row: results.Range(
                    f"B{start_row}:L{end_row}"
                )
            )
            actual_matrix = tuple(
                tuple(_as_text(value) for value in row)
                for row in _matrix_values(retry_com_call(lambda: result_range.Value2))
            )
            _check(
                report,
                code="result_channel",
                target=f"Resultados!B{start_row}:L{end_row}",
                expected=expected_matrix,
                actual=actual_matrix,
                passed=actual_matrix == expected_matrix,
            )

        for row, note in zip(legacy._NOTE_TARGET_ROWS, payload.get("notes") or []):
            note_cell = retry_com_call(lambda row=row: results.Range(f"B{row}"))
            actual = _as_text(retry_com_call(lambda: note_cell.Value2))
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


def _remove_with_retry(path: Path, *, attempts: int = 12) -> None:
    if not path.exists():
        return
    for attempt in range(attempts):
        try:
            path.unlink(missing_ok=True)
            return
        except PermissionError:
            if attempt == attempts - 1:
                raise RuntimeError(
                    f"Could not remove locked workbook: {path}. Close any visible Excel window "
                    "and pause OneDrive synchronization before retrying."
                )
            time.sleep(0.25 + attempt * 0.1)


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
    legacy._require_complete_payload(payload)

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
        _remove_with_retry(output_path)

    shutil.copy2(template_path.resolve(), output_path)
    plan = legacy.build_scalar_write_plan(payload)

    try:
        with open_excel_workbook(output_path, read_only=False) as (excel, workbook):
            if bool(retry_com_call(lambda: workbook.ReadOnly)):
                raise RuntimeError("Copied workbook unexpectedly opened as read-only.")
            _write_scalar_fields(workbook, plan)
            info_sheet = retry_com_call(lambda: workbook.Worksheets("Información"))
            result_sheet = retry_com_call(lambda: workbook.Worksheets("Resultados"))
            _write_method(info_sheet, str(payload["calibration_method"]))
            _write_traceability(info_sheet, list(payload["traceability"]))
            _write_results(result_sheet, payload)
            retry_com_call(excel.CalculateFullRebuild, attempts=20, initial_delay=0.25)
            retry_com_call(workbook.Save, attempts=20, initial_delay=0.25)
    except Exception as exc:
        try:
            _remove_with_retry(output_path)
        except RuntimeError as cleanup_exc:
            raise RuntimeError(f"{exc}\nCleanup also failed: {cleanup_exc}") from exc
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
            "information_print_area": f"B1:H{legacy._INFO_PRINT_END_ROW}",
            "results_print_area": f"B1:L{legacy._RESULTS_PRINT_END_ROW}",
            "result_page_breaks_before_rows": list(legacy._RESULT_PAGE_BREAKS),
            "channel_layout": {
                channel: {
                    "header_row": header,
                    "data_start_row": start,
                    "data_end_row": end,
                }
                for channel, (header, start, end) in legacy._CHANNEL_LAYOUT.items()
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
