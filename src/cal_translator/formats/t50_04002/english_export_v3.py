from __future__ import annotations

from pathlib import Path
from typing import Any

from cal_translator.formats.t50_04002 import english_export_v2 as v2


def try_enable_shape_printing(shape: Any) -> bool:
    """Enable shape printing when Excel exposes PrintObject as writable.

    Text boxes created through Shapes.AddTextbox are printed by default. Some
    Excel/pywin32 combinations expose Shape.PrintObject as read-only, so failure
    to assign this optional property must not abort certificate generation.
    """
    try:
        v2.set_com_property(shape, "PrintObject", True)
    except AttributeError:
        return False
    return True


def _create_laboratory_text_box(sheet: Any, spec: dict[str, Any]) -> dict[str, Any]:
    for row, height in spec["row_heights"].items():
        row_object = v2.retry_com_call(lambda row=row: sheet.Rows(row))
        v2.set_com_property(row_object, "RowHeight", height)

    anchor = v2.retry_com_call(lambda: sheet.Range(spec["anchor"]))
    left = float(v2.retry_com_call(lambda: anchor.Left))
    top = float(v2.retry_com_call(lambda: anchor.Top))
    removed_existing = v2._delete_named_shape(sheet, spec["name"])

    shapes = v2.retry_com_call(lambda: sheet.Shapes)
    shape = v2.retry_com_call(
        lambda: shapes.AddTextbox(
            v2._MSO_TEXT_ORIENTATION_HORIZONTAL,
            left,
            top,
            spec["width"],
            spec["height"],
        )
    )
    v2.set_com_property(shape, "Name", spec["name"])
    v2.set_com_property(shape, "Placement", v2._XL_MOVE_AND_SIZE)
    print_object_configured = try_enable_shape_printing(shape)

    fill = v2.retry_com_call(lambda: shape.Fill)
    line = v2.retry_com_call(lambda: shape.Line)
    v2.set_com_property(fill, "Visible", v2._MSO_FALSE)
    v2.set_com_property(line, "Visible", v2._MSO_FALSE)

    text_frame = v2.retry_com_call(lambda: shape.TextFrame2)
    v2.set_com_property(text_frame, "MarginLeft", 0)
    v2.set_com_property(text_frame, "MarginRight", 0)
    v2.set_com_property(text_frame, "MarginTop", 0)
    v2.set_com_property(text_frame, "MarginBottom", 0)
    v2.set_com_property(text_frame, "WordWrap", True)

    text_range = v2.retry_com_call(lambda: text_frame.TextRange)
    v2.set_com_property(text_range, "Text", spec["text"])
    font = v2.retry_com_call(lambda: text_range.Font)
    v2.set_com_property(font, "Name", spec["font_name"])
    v2.set_com_property(font, "Size", spec["font_size"])

    return {
        "shape_name": spec["name"],
        "anchor": spec["anchor"],
        "text": spec["text"],
        "width": spec["width"],
        "height": spec["height"],
        "font_name": spec["font_name"],
        "font_size": spec["font_size"],
        "row_heights": spec["row_heights"],
        "removed_existing_shape": removed_existing,
        "print_object_configured": print_object_configured,
        "print_behavior": (
            "explicit" if print_object_configured else "excel_default"
        ),
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
    """Run v2 export while using a version-tolerant text-box creator."""
    original_creator = v2._create_laboratory_text_box
    v2._create_laboratory_text_box = _create_laboratory_text_box
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
        v2._create_laboratory_text_box = original_creator
