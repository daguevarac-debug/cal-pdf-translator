from __future__ import annotations

import copy
import json
from importlib import resources
from pathlib import Path
from typing import Any

from cal_translator.excel.com_backend import retry_com_call, set_com_property
from cal_translator.formats.t50_04002 import english_export as legacy

_RESULTS_PRINT_END_ROW = 102
_OPTIONAL_RESULT_RANGE = "I10:L83"
_XL_FALSE = False
_XL_CENTER = -4108


def _matrix(value: Any) -> tuple[tuple[Any, ...], ...]:
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return tuple(tuple(row) for row in value)
        return (tuple(value),)
    return ((value,),)


def _load_overrides() -> dict[str, Any]:
    try:
        import yaml
    except ImportError as exc:
        raise RuntimeError(
            "PyYAML is required. Install the project with: python -m pip install -e ."
        ) from exc

    resource = resources.files(__package__).joinpath(
        "american_english_export_overrides.yaml"
    )
    with resource.open("r", encoding="utf-8") as handle:
        overrides = yaml.safe_load(handle)
    if not isinstance(overrides, dict):
        raise RuntimeError("Invalid controlled English export override profile.")
    return overrides


def merge_controlled_profile(
    profile: dict[str, Any], overrides: dict[str, Any]
) -> dict[str, Any]:
    merged = copy.deepcopy(profile)
    merged["profile_version"] = overrides.get(
        "profile_version", merged.get("profile_version")
    )

    validation = merged.setdefault("pdf_validation", {})
    override_validation = overrides.get("pdf_validation") or {}
    for key in ("required_phrases", "forbidden_phrases"):
        values = list(validation.get(key) or [])
        for item in override_validation.get(key) or []:
            if item not in values:
                values.append(item)
        validation[key] = values
    return merged


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


def _delete_vertical_page_breaks(sheet: Any) -> int:
    breaks = retry_com_call(lambda: sheet.VPageBreaks)
    count = int(retry_com_call(lambda: breaks.Count))
    deleted = 0
    for _ in range(count):
        page_break = retry_com_call(lambda: breaks(1))
        retry_com_call(page_break.Delete)
        deleted += 1
    return deleted


def configure_pdf_print_layout(workbook: Any) -> dict[str, Any]:
    """Configure native Excel pagination without dropping non-empty result data."""
    results = legacy._get_sheet(workbook, "Resultados")
    optional_range = retry_com_call(lambda: results.Range(_OPTIONAL_RESULT_RANGE))
    optional_values = retry_com_call(lambda: optional_range.Value2)
    layout = select_results_print_layout(optional_values)

    page_setup = retry_com_call(lambda: results.PageSetup)
    set_com_property(page_setup, "PrintArea", layout["print_area"])

    deleted_breaks = 0
    if layout["fit_to_one_page_wide"]:
        deleted_breaks = _delete_vertical_page_breaks(results)
        set_com_property(page_setup, "Zoom", _XL_FALSE)
        set_com_property(page_setup, "FitToPagesWide", 1)
        set_com_property(page_setup, "FitToPagesTall", _XL_FALSE)

    layout["vertical_page_breaks_deleted"] = deleted_breaks
    return layout


def information_left_header_text(overrides: dict[str, Any]) -> str:
    header = overrides.get("information_left_header")
    if not isinstance(header, str) or "&G" not in header:
        raise RuntimeError(
            "The controlled information-page header must be a string containing the &G logo token."
        )
    return header.rstrip()


def traceability_presentation(overrides: dict[str, Any]) -> dict[str, Any]:
    layout = overrides.get("traceability_layout") or {}
    heights = layout.get("row_heights") or {}
    normalized_heights = {int(row): float(height) for row, height in heights.items()}
    expected_rows = set(range(69, 74))
    if set(normalized_heights) != expected_rows:
        raise RuntimeError(
            "Controlled traceability layout must define row heights for rows 69 through 73."
        )
    return {
        "equipment_font_size": float(layout.get("equipment_font_size", 8.5)),
        "row_heights": normalized_heights,
    }


def title_presentation(overrides: dict[str, Any]) -> dict[str, str]:
    title = overrides.get("title_layout") or {}
    certificate_label = str(title.get("certificate_number_label") or "").strip()
    duplicate_format = str(title.get("duplicate_certificate_number_format") or "").strip()
    if not certificate_label or not duplicate_format:
        raise RuntimeError(
            "Controlled title layout must define the certificate label and duplicate-value format."
        )
    return {
        "certificate_number_label": certificate_label,
        "duplicate_certificate_number_format": duplicate_format,
    }


def configure_information_presentation(
    workbook: Any, overrides: dict[str, Any]
) -> dict[str, Any]:
    """Restore the laboratory header and make first-page tables and title legible."""
    information = legacy._get_sheet(workbook, "Información")
    page_setup = retry_com_call(lambda: information.PageSetup)
    left_header = information_left_header_text(overrides)
    set_com_property(page_setup, "LeftHeader", left_header)

    title_layout = title_presentation(overrides)
    certificate_label = retry_com_call(lambda: information.Range("B4"))
    set_com_property(
        certificate_label,
        "Value2",
        title_layout["certificate_number_label"],
    )
    duplicate_certificate = retry_com_call(lambda: information.Range("B6"))
    set_com_property(
        duplicate_certificate,
        "NumberFormat",
        title_layout["duplicate_certificate_number_format"],
    )

    layout = traceability_presentation(overrides)
    traceability_range = retry_com_call(lambda: information.Range("B69:H73"))
    set_com_property(traceability_range, "WrapText", True)
    set_com_property(traceability_range, "VerticalAlignment", _XL_CENTER)

    equipment_range = retry_com_call(lambda: information.Range("B69:C73"))
    equipment_font = retry_com_call(lambda: equipment_range.Font)
    set_com_property(equipment_font, "Size", layout["equipment_font_size"])

    for row, height in layout["row_heights"].items():
        row_object = retry_com_call(lambda row=row: information.Rows(row))
        set_com_property(row_object, "RowHeight", height)

    return {
        "left_header": left_header,
        "certificate_number_label": title_layout["certificate_number_label"],
        "duplicate_certificate_number_format": title_layout[
            "duplicate_certificate_number_format"
        ],
        "equipment_font_size": layout["equipment_font_size"],
        "row_heights": layout["row_heights"],
    }


def configure_controlled_metadata(
    workbook: Any, overrides: dict[str, Any]
) -> dict[str, Any]:
    """Translate workbook metadata that is stored outside ordinary cell values."""
    information = legacy._get_sheet(workbook, "Información")
    internal_code = retry_com_call(lambda: information.Range("B7"))
    number_format = str(overrides["internal_code_number_format"])
    set_com_property(internal_code, "NumberFormat", number_format)

    right_footer = str(overrides["right_footer"])
    footer_sheets: list[str] = []
    for name in ("Información", "Resultados"):
        sheet = legacy._get_sheet(workbook, name)
        page_setup = retry_com_call(lambda sheet=sheet: sheet.PageSetup)
        set_com_property(page_setup, "RightFooter", right_footer)
        footer_sheets.append(name)

    return {
        "internal_code_number_format": number_format,
        "right_footer": right_footer,
        "footer_sheets": footer_sheets,
    }


def translate_and_export(
    source_workbook: Path,
    data_path: Path,
    output_workbook: Path,
    output_pdf: Path,
    *,
    format_id: str = legacy.FORMAT_ID,
    overwrite: bool = False,
) -> tuple[dict[str, Any], Path]:
    """Run the controlled English export with corrected pagination and presentation."""
    export_layout: dict[str, Any] = {}
    metadata_changes: dict[str, Any] = {}
    information_presentation: dict[str, Any] = {}
    overrides = _load_overrides()

    original_export = legacy._export_certificate_pdf
    original_profile_loader = legacy._load_profile

    def load_controlled_profile() -> dict[str, Any]:
        return merge_controlled_profile(original_profile_loader(), overrides)

    def export_with_controlled_layout(workbook: Any, pdf_path: Path) -> None:
        metadata_changes.update(configure_controlled_metadata(workbook, overrides))
        information_presentation.update(
            configure_information_presentation(workbook, overrides)
        )
        export_layout.update(configure_pdf_print_layout(workbook))
        retry_com_call(workbook.Save, attempts=20, initial_delay=0.25)
        original_export(workbook, pdf_path)

    legacy._load_profile = load_controlled_profile
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
        legacy._load_profile = original_profile_loader
        legacy._export_certificate_pdf = original_export

    report["pdf_layout"] = export_layout
    report["controlled_metadata"] = metadata_changes
    report["information_presentation"] = information_presentation
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    return report, report_path
