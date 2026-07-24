from __future__ import annotations

from pathlib import Path
from typing import Any

from cal_translator.excel.com_backend import retry_com_call, set_com_property
from cal_translator.formats.t50_04002 import temperature_workbook_writer as legacy
from cal_translator.formats.t50_04002 import workbook_writer as common

_ORIGINAL_BUILD_TEMPERATURE_WORKBOOK = legacy.build_temperature_workbook
_ORIGINAL_INFO_PRINT_END = legacy._INFO_PRINT_END

_TRACE_COLUMNS: tuple[tuple[str, str], ...] = (
    ("B", "equipment"),
    ("D", "type"),
    ("E", "internal_number"),
    ("F", "calibrated_by"),
    ("G", "certificate_number"),
    ("H", "calibration_date"),
)


def _merge_area(target: Any) -> Any:
    """Return the complete merged area, or the original range when unmerged."""
    is_merged = bool(retry_com_call(lambda: target.MergeCells))
    return retry_com_call(lambda: target.MergeArea) if is_merged else target


def _clear_and_write_logical_cell(sheet: Any, address: str, value: Any) -> None:
    """Safely replace one logical cell without changing its merge structure."""
    target = retry_com_call(lambda: sheet.Range(address))
    area = _merge_area(target)
    retry_com_call(area.ClearContents)
    top_left = retry_com_call(lambda: area.Cells(1, 1))
    set_com_property(top_left, "Value2", legacy._as_text(value))


def _write_traceability(info_sheet: Any, entries: list[dict[str, Any]]) -> None:
    """Use the three traceability rows already present in the controlled template."""
    for offset, entry in enumerate(entries):
        row = legacy._TRACE_START + offset
        equipment_area = retry_com_call(lambda row=row: info_sheet.Range(f"B{row}:C{row}"))
        if not bool(retry_com_call(lambda: equipment_area.MergeCells)):
            retry_com_call(equipment_area.Merge)

        set_com_property(equipment_area, "NumberFormat", "@")
        set_com_property(equipment_area, "WrapText", True)
        set_com_property(equipment_area, "VerticalAlignment", common._XL_CENTER)
        _clear_and_write_logical_cell(info_sheet, f"B{row}", entry.get("equipment"))

        for column, field in _TRACE_COLUMNS[1:]:
            cell = retry_com_call(lambda column=column, row=row: info_sheet.Range(f"{column}{row}"))
            set_com_property(cell, "NumberFormat", "@")
            set_com_property(cell, "WrapText", True)
            set_com_property(cell, "VerticalAlignment", common._XL_CENTER)
            _clear_and_write_logical_cell(info_sheet, f"{column}{row}", entry.get(field))

        row_object = retry_com_call(lambda row=row: info_sheet.Rows(row))
        set_com_property(row_object, "RowHeight", 22)

    page_setup = retry_com_call(lambda: info_sheet.PageSetup)
    set_com_property(page_setup, "PrintArea", "$B$1:$H$77")


def _write_results(result_sheet: Any, payload: dict[str, Any]) -> None:
    """Write temperature results and notes without clearing partial merged cells."""
    metadata = payload.get("result_metadata") or {}
    legacy._configure_merged_title(
        result_sheet,
        legacy._RESULT_TITLE_ROW,
        str(metadata.get("measurement_title") or "Medición de Temperatura"),
    )

    labels = {
        f"B{legacy._RESULT_ACCURACY_ROW}": "Exactitud",
        f"C{legacy._RESULT_ACCURACY_ROW}": legacy._as_text(metadata.get("accuracy")),
        f"B{legacy._RESULT_RANGE_ROW}": "Desde",
        f"C{legacy._RESULT_RANGE_ROW}": legacy._as_text(metadata.get("from")),
        f"D{legacy._RESULT_RANGE_ROW}": "Hasta",
        f"E{legacy._RESULT_RANGE_ROW}": legacy._as_text(metadata.get("to")),
    }
    for address, value in labels.items():
        _clear_and_write_logical_cell(result_sheet, address, value)

    for address in (
        f"B{legacy._RESULT_ACCURACY_ROW}",
        f"B{legacy._RESULT_RANGE_ROW}",
        f"D{legacy._RESULT_RANGE_ROW}",
    ):
        font = retry_com_call(lambda address=address: result_sheet.Range(address).Font)
        set_com_property(font, "Bold", True)

    for address in (f"C{legacy._RESULT_RANGE_ROW}", f"E{legacy._RESULT_RANGE_ROW}"):
        font = retry_com_call(lambda address=address: result_sheet.Range(address).Font)
        set_com_property(font, "Color", 255)
        set_com_property(font, "Bold", True)

    target = retry_com_call(
        lambda: result_sheet.Range(f"C{legacy._RESULT_START}:H{legacy._RESULT_END}")
    )
    retry_com_call(target.ClearContents)
    set_com_property(target, "NumberFormat", "@")
    set_com_property(target, "Value2", legacy.temperature_result_matrix(payload))
    set_com_property(target, "HorizontalAlignment", common._XL_CENTER)
    set_com_property(target, "VerticalAlignment", common._XL_CENTER)
    set_com_property(target, "WrapText", False)
    rows = retry_com_call(
        lambda: result_sheet.Rows(f"{legacy._RESULT_START}:{legacy._RESULT_END}")
    )
    set_com_property(rows, "RowHeight", 14.5)

    for row, note in zip(legacy._NOTE_ROWS, payload.get("notes") or []):
        _clear_and_write_logical_cell(result_sheet, f"B{row}", str(note))

    page_setup = retry_com_call(lambda: result_sheet.PageSetup)
    set_com_property(page_setup, "PrintArea", f"$B$1:$H${legacy._RESULT_PRINT_END}")
    set_com_property(page_setup, "PrintTitleRows", "")
    retry_com_call(result_sheet.ResetAllPageBreaks)


def build_temperature_workbook(
    template_path: Path,
    data_path: Path,
    output_path: Path,
    *,
    format_id: str = legacy.FORMAT_ID,
    overwrite: bool = False,
):
    """Run the temperature writer with merge-safe writes and no row insertion."""
    original_traceability = legacy._write_traceability
    original_results = legacy._write_results
    original_info_end = legacy._INFO_PRINT_END
    legacy._write_traceability = _write_traceability
    legacy._write_results = _write_results
    legacy._INFO_PRINT_END = 77
    try:
        return _ORIGINAL_BUILD_TEMPERATURE_WORKBOOK(
            template_path,
            data_path,
            output_path,
            format_id=format_id,
            overwrite=overwrite,
        )
    finally:
        legacy._write_traceability = original_traceability
        legacy._write_results = original_results
        legacy._INFO_PRINT_END = original_info_end


def install_patch() -> None:
    """Expose the merge-safe writer through the module imported by the batch CLI."""
    legacy.build_temperature_workbook = build_temperature_workbook
