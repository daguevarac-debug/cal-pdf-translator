from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from cal_translator.excel.com_backend import retry_com_call, set_com_property
from cal_translator.formats.t50_04002 import english_export as legacy

_RESULTS_PRINT_END_ROW = 102
_OPTIONAL_RESULT_RANGE = "I10:L83"
_XL_FALSE = False


def _matrix(value: Any) -> tuple[tuple[Any, ...], ...]:
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return tuple(tuple(row) for row in value)
        return (tuple(value),)
    return ((value,),)


def optional_result_columns_are_empty(value: Any) -> bool:
    """Return true when CMC, Pass/Fail, TUR and TAR contain no published data."""
    return all(
        cell is None or str(cell).strip() == ""
        for row in _matrix(value)
        for cell in row
    )


def select_results_print_layout(optional_values: Any) -> dict[str, Any]:
    optional_empty = optional_result_columns_are_empty(optional_values)
    return {
        "optional_columns_empty": optional_empty,
        "optional_columns_included": not optional_empty,
        "print_area": (
            f"$B$1:$H${_RESULTS_PRINT_END_ROW}"
            if optional_empty
            else f"$B$1:$L${_RESULTS_PRINT_END_ROW}"
        ),
        "fit_to_one_page_wide": not optional_empty,
    }


def configure_pdf_print_layout(workbook: Any) -> dict[str, Any]:
    """Configure native Excel pagination without dropping non-empty result data."""
    results = legacy._get_sheet(workbook, "Resultados")
    optional_range = retry_com_call(lambda: results.Range(_OPTIONAL_RESULT_RANGE))
    optional_values = retry_com_call(lambda: optional_range.Value2)
    layout = select_results_print_layout(optional_values)

    page_setup = retry_com_call(lambda: results.PageSetup)
    set_com_property(page_setup, "PrintArea", layout["print_area"])

    if layout["fit_to_one_page_wide"]:
        set_com_property(page_setup, "Zoom", _XL_FALSE)
        set_com_property(page_setup, "FitToPagesWide", 1)
        set_com_property(page_setup, "FitToPagesTall", _XL_FALSE)

    return layout


def translate_and_export(
    source_workbook: Path,
    data_path: Path,
    output_workbook: Path,
    output_pdf: Path,
    *,
    format_id: str = legacy.FORMAT_ID,
    overwrite: bool = False,
) -> tuple[dict[str, Any], Path]:
    """Run the controlled English export with corrected horizontal pagination."""
    export_layout: dict[str, Any] = {}
    original_export = legacy._export_certificate_pdf

    def export_with_controlled_layout(workbook: Any, pdf_path: Path) -> None:
        export_layout.update(configure_pdf_print_layout(workbook))
        retry_com_call(workbook.Save, attempts=20, initial_delay=0.25)
        original_export(workbook, pdf_path)

    legacy._export_certificate_pdf = export_with_controlled_layout
    try:
        report, report_path = legacy.translate_and_export(
            source_workbook,
            data_path,
            output_workbook,
            output_pdf,
            format_id=format_id,
            overwrite=overwrite,
        )
    finally:
        legacy._export_certificate_pdf = original_export

    report["pdf_layout"] = export_layout
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    return report, report_path
