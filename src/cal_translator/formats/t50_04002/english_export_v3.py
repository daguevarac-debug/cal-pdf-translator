from __future__ import annotations

from pathlib import Path
from typing import Any

from cal_translator.formats.t50_04002 import english_export_v2 as v2

_XL_LEFT = -4131
_XL_TOP = -4160
_LOGO_ONLY_HEADER = "&G"
_LEGACY_SHAPE_NAME = "CAL_English_Laboratory_Block"


def laboratory_cell_specs(overrides: dict[str, Any]) -> list[dict[str, Any]]:
    """Return the controlled cell blocks used below the first-page logo."""
    blocks = overrides.get("information_cells") or []
    if not isinstance(blocks, list) or len(blocks) != 2:
        raise RuntimeError(
            "Controlled information cells must define exactly two cell blocks."
        )

    specs: list[dict[str, Any]] = []
    for index, block in enumerate(blocks):
        if not isinstance(block, dict):
            raise RuntimeError(f"Information cell block {index} must be an object.")
        address = str(block.get("range") or "").strip()
        text = str(block.get("text") or "").strip()
        row = int(block.get("row") or 0)
        height = float(block.get("row_height") or 0)
        if not address or not text or row <= 0 or height <= 0:
            raise RuntimeError(
                f"Information cell block {index} must define range, text, row and row_height."
            )
        specs.append(
            {
                "range": address,
                "text": text,
                "row": row,
                "row_height": height,
                "font_name": str(block.get("font_name") or "Arial"),
                "font_size": float(block.get("font_size") or 8),
                "bold": bool(block.get("bold", False)),
            }
        )
    return specs


def _set_logo_only_header(sheet: Any) -> bool:
    """Keep the existing header image while removing its old text caption."""
    page_setup = v2.retry_com_call(lambda: sheet.PageSetup)
    try:
        v2.set_com_property(page_setup, "LeftHeader", _LOGO_ONLY_HEADER)
    except Exception:  # noqa: BLE001 - tolerate Office builds that lock picture headers
        return False
    return True


def _write_merged_cell_block(sheet: Any, spec: dict[str, Any]) -> dict[str, Any]:
    """Write one laboratory identity block into real worksheet cells."""
    target = v2.retry_com_call(lambda: sheet.Range(spec["range"]))
    try:
        v2.retry_com_call(target.UnMerge)
    except Exception:  # noqa: BLE001 - range may already be unmerged
        pass
    v2.retry_com_call(target.ClearContents)
    v2.retry_com_call(target.Merge)

    top_left = v2.retry_com_call(lambda: target.Cells(1, 1))
    v2.set_com_property(top_left, "Value2", spec["text"])
    v2.set_com_property(target, "WrapText", True)
    v2.set_com_property(target, "ShrinkToFit", False)
    v2.set_com_property(target, "HorizontalAlignment", _XL_LEFT)
    v2.set_com_property(target, "VerticalAlignment", _XL_TOP)

    font = v2.retry_com_call(lambda: target.Font)
    v2.set_com_property(font, "Name", spec["font_name"])
    v2.set_com_property(font, "Size", spec["font_size"])
    v2.set_com_property(font, "Bold", spec["bold"])

    row_object = v2.retry_com_call(lambda: sheet.Rows(spec["row"]))
    v2.set_com_property(row_object, "RowHeight", spec["row_height"])

    return dict(spec)


def configure_information_presentation(
    workbook: Any, overrides: dict[str, Any]
) -> dict[str, Any]:
    """Use worksheet cells for the first-page identity and preserve normal layout."""
    information = v2.legacy._get_sheet(workbook, "Información")
    removed_legacy_shape = v2._delete_named_shape(information, _LEGACY_SHAPE_NAME)
    logo_only_header = _set_logo_only_header(information)

    cell_blocks = [
        _write_merged_cell_block(information, spec)
        for spec in laboratory_cell_specs(overrides)
    ]

    title_layout = v2.title_presentation(overrides)
    certificate_label = v2.retry_com_call(lambda: information.Range("B4"))
    v2.set_com_property(
        certificate_label,
        "Value2",
        title_layout["certificate_number_label"],
    )
    duplicate_certificate = v2.retry_com_call(lambda: information.Range("B6"))
    v2.set_com_property(
        duplicate_certificate,
        "NumberFormat",
        title_layout["duplicate_certificate_number_format"],
    )

    layout = v2.traceability_presentation(overrides)
    traceability_range = v2.retry_com_call(lambda: information.Range("B69:H73"))
    v2.set_com_property(traceability_range, "WrapText", True)
    v2.set_com_property(traceability_range, "VerticalAlignment", v2._XL_CENTER)

    equipment_range = v2.retry_com_call(lambda: information.Range("B69:C73"))
    equipment_font = v2.retry_com_call(lambda: equipment_range.Font)
    v2.set_com_property(equipment_font, "Size", layout["equipment_font_size"])

    for row, height in layout["row_heights"].items():
        row_object = v2.retry_com_call(lambda row=row: information.Rows(row))
        v2.set_com_property(row_object, "RowHeight", height)

    return {
        "presentation_mode": "worksheet_cells",
        "removed_legacy_shape": removed_legacy_shape,
        "logo_only_header_configured": logo_only_header,
        "information_cells": cell_blocks,
        "certificate_number_label": title_layout["certificate_number_label"],
        "duplicate_certificate_number_format": title_layout[
            "duplicate_certificate_number_format"
        ],
        "equipment_font_size": layout["equipment_font_size"],
        "row_heights": layout["row_heights"],
    }


def translate_and_export(
    source_workbook: Path,
    data_path: Path,
    output_workbook: Path,
    output_pdf: Path,
    *,
    format_id: str = v2.legacy.FORMAT_ID,
    overwrite: bool = False,
) -> tuple[dict[str, Any], Path]:
    """Run v2 export while rendering the first-page identity in cells."""
    original_presentation = v2.configure_information_presentation
    v2.configure_information_presentation = configure_information_presentation
    try:
        return v2.translate_and_export(
            source_workbook,
            data_path,
            output_workbook,
            output_pdf,
            format_id=format_id,
            overwrite=overwrite,
        )
    finally:
        v2.configure_information_presentation = original_presentation
