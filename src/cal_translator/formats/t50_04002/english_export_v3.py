from __future__ import annotations

from pathlib import Path
from typing import Any

from cal_translator.formats.t50_04002 import english_export_v2 as v2

_MSO_TRUE = -1
_WHITE_RGB = 16777215


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


def laboratory_text_box_spec(overrides: dict[str, Any]) -> dict[str, Any]:
    """Build the controlled first-page laboratory block layout."""
    block = overrides.get("information_text_box") or {}
    required_text = str(block.get("text") or "").strip()
    name = str(block.get("name") or "").strip()
    anchor = str(block.get("anchor") or "").strip()
    row_heights = {
        int(row): float(height)
        for row, height in (block.get("row_heights") or {}).items()
    }
    if not required_text or not name or not anchor:
        raise RuntimeError(
            "Controlled information text box must define name, anchor and text."
        )
    if set(row_heights) != {1, 2}:
        raise RuntimeError(
            "Controlled information text box must define row heights for rows 1 and 2."
        )
    return {
        "name": name,
        "anchor": anchor,
        "text": required_text,
        "width": float(block.get("width", 360)),
        "height": float(block.get("height", 58)),
        "left_offset": float(block.get("left_offset", 0)),
        "top_offset": float(block.get("top_offset", 0)),
        "font_name": str(block.get("font_name") or "Arial"),
        "font_size": float(block.get("font_size", 7.5)),
        "text_margin_left": float(block.get("text_margin_left", 0)),
        "text_margin_right": float(block.get("text_margin_right", 0)),
        "text_margin_top": float(block.get("text_margin_top", 0)),
        "text_margin_bottom": float(block.get("text_margin_bottom", 0)),
        "white_fill": bool(block.get("white_fill", False)),
        "row_heights": row_heights,
    }


def _configure_shape_fill(shape: Any, *, white_fill: bool) -> None:
    fill = v2.retry_com_call(lambda: shape.Fill)
    line = v2.retry_com_call(lambda: shape.Line)
    v2.set_com_property(line, "Visible", v2._MSO_FALSE)

    if not white_fill:
        v2.set_com_property(fill, "Visible", v2._MSO_FALSE)
        return

    v2.set_com_property(fill, "Visible", _MSO_TRUE)
    try:
        v2.retry_com_call(fill.Solid)
    except Exception:  # noqa: BLE001 - Solid is unavailable in some Excel builds
        pass
    fore_color = v2.retry_com_call(lambda: fill.ForeColor)
    v2.set_com_property(fore_color, "RGB", _WHITE_RGB)
    try:
        v2.set_com_property(fill, "Transparency", 0.0)
    except Exception:  # noqa: BLE001 - optional across Office versions
        pass


def _create_laboratory_text_box(sheet: Any, spec: dict[str, Any]) -> dict[str, Any]:
    for row, height in spec["row_heights"].items():
        row_object = v2.retry_com_call(lambda row=row: sheet.Rows(row))
        v2.set_com_property(row_object, "RowHeight", height)

    anchor = v2.retry_com_call(lambda: sheet.Range(spec["anchor"]))
    left = float(v2.retry_com_call(lambda: anchor.Left)) + spec["left_offset"]
    top = float(v2.retry_com_call(lambda: anchor.Top)) + spec["top_offset"]
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
    _configure_shape_fill(shape, white_fill=spec["white_fill"])

    text_frame = v2.retry_com_call(lambda: shape.TextFrame2)
    v2.set_com_property(text_frame, "MarginLeft", spec["text_margin_left"])
    v2.set_com_property(text_frame, "MarginRight", spec["text_margin_right"])
    v2.set_com_property(text_frame, "MarginTop", spec["text_margin_top"])
    v2.set_com_property(text_frame, "MarginBottom", spec["text_margin_bottom"])
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
        "left_offset": spec["left_offset"],
        "top_offset": spec["top_offset"],
        "font_name": spec["font_name"],
        "font_size": spec["font_size"],
        "white_fill": spec["white_fill"],
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
    """Run v2 export with the controlled, version-tolerant first-page block."""
    original_spec = v2.laboratory_text_box_spec
    original_creator = v2._create_laboratory_text_box
    v2.laboratory_text_box_spec = laboratory_text_box_spec
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
        v2.laboratory_text_box_spec = original_spec
        v2._create_laboratory_text_box = original_creator
