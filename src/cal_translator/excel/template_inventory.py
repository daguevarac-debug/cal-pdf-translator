from __future__ import annotations

import json
import logging
import os
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

LOGGER = logging.getLogger(__name__)

_EDGE_INDEXES = {
    "left": 7,
    "top": 8,
    "bottom": 9,
    "right": 10,
}

_ROLE_PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    ("certificate_number", re.compile(r"certificad[oa].*calibraci[oó]n|calibration certificate", re.I)),
    ("document_code", re.compile(r"c[oó]digo.*documento|document code", re.I)),
    ("equipment_description", re.compile(r"descripci[oó]n|description", re.I)),
    ("equipment_manufacturer", re.compile(r"fabricante|manufacturer", re.I)),
    ("equipment_serial_number", re.compile(r"n[uú]mero.*serie|serial number", re.I)),
    ("equipment_model", re.compile(r"modelo|model", re.I)),
    ("customer", re.compile(r"cliente|customer", re.I)),
    ("calibration_method", re.compile(r"m[eé]todo.*calibraci[oó]n|calibration method", re.I)),
    ("environment", re.compile(r"temperatura|temperature|humedad|humidity", re.I)),
    ("date", re.compile(r"fecha|date", re.I)),
    ("measurement_uncertainty", re.compile(r"incertidumbre|uncertainty", re.I)),
    ("traceability", re.compile(r"trazabilidad|traceability", re.I)),
    ("results", re.compile(r"resultados|results|rango|range|sesgo|bias", re.I)),
    ("notes", re.compile(r"notas|notes", re.I)),
    ("observations", re.compile(r"observaciones|observations", re.I)),
]


def _safe_get(getter: Any, default: Any = None) -> Any:
    try:
        return getter()
    except Exception:  # noqa: BLE001
        return default


def _json_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value)


def _address(obj: Any) -> str | None:
    return _safe_get(lambda: str(obj.Address(False, False)))


def _cell_type(value: Any, has_formula: bool) -> str:
    if has_formula:
        return "formula"
    if value is None or value == "":
        return "blank"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    return "text"


def _probable_role(sheet_name: str, text: str, value: Any, has_formula: bool) -> str:
    normalized = re.sub(r"\s+", " ", text or "").strip()
    if has_formula:
        return "formula"
    if normalized in {"#N/A", "#N/D", "#NA"}:
        return "variable_placeholder"
    for role, pattern in _ROLE_PATTERNS:
        if pattern.search(normalized):
            return role
    if sheet_name.casefold() == "resultados":
        if isinstance(value, (int, float)):
            return "result_value"
        return "results_content"
    return "static_or_variable_content"


def _border_snapshot(cell: Any) -> dict[str, Any]:
    borders: dict[str, Any] = {}
    for name, index in _EDGE_INDEXES.items():
        border = _safe_get(lambda i=index: cell.Borders(i))
        if border is None:
            continue
        line_style = _safe_get(lambda b=border: b.LineStyle)
        if line_style in (None, 0):
            continue
        borders[name] = {
            "line_style": _json_value(line_style),
            "weight": _json_value(_safe_get(lambda b=border: b.Weight)),
            "color": _json_value(_safe_get(lambda b=border: b.Color)),
        }
    return borders


def _cell_snapshot(cell: Any, sheet_name: str) -> dict[str, Any] | None:
    has_formula = bool(_safe_get(lambda: cell.HasFormula, False))
    value = _safe_get(lambda: cell.Value2)
    text = str(_safe_get(lambda: cell.Text, "") or "")
    comment = _safe_get(lambda: cell.Comment)
    comment_text = _safe_get(lambda: comment.Text(), None) if comment is not None else None

    if value in (None, "") and not has_formula and not comment_text:
        return None

    merge_area = None
    if bool(_safe_get(lambda: cell.MergeCells, False)):
        merge_area = _safe_get(lambda: str(cell.MergeArea.Address(False, False)))

    formula = _safe_get(lambda: cell.Formula, None) if has_formula else None
    return {
        "address": _address(cell),
        "value": _json_value(value),
        "display_text": text,
        "cell_type": _cell_type(value, has_formula),
        "formula": _json_value(formula),
        "number_format": _json_value(_safe_get(lambda: cell.NumberFormatLocal)),
        "style": _json_value(_safe_get(lambda: cell.Style)),
        "merged_range": merge_area,
        "probable_role": _probable_role(sheet_name, text, value, has_formula),
        "alignment": {
            "horizontal": _json_value(_safe_get(lambda: cell.HorizontalAlignment)),
            "vertical": _json_value(_safe_get(lambda: cell.VerticalAlignment)),
            "wrap_text": bool(_safe_get(lambda: cell.WrapText, False)),
            "orientation": _json_value(_safe_get(lambda: cell.Orientation)),
        },
        "font": {
            "name": _json_value(_safe_get(lambda: cell.Font.Name)),
            "size": _json_value(_safe_get(lambda: cell.Font.Size)),
            "bold": bool(_safe_get(lambda: cell.Font.Bold, False)),
            "italic": bool(_safe_get(lambda: cell.Font.Italic, False)),
            "color": _json_value(_safe_get(lambda: cell.Font.Color)),
        },
        "fill_color": _json_value(_safe_get(lambda: cell.Interior.Color)),
        "borders": _border_snapshot(cell),
        "comment": _json_value(comment_text),
    }


def _collect_cells(sheet: Any, used_range: Any) -> tuple[list[dict[str, Any]], list[str]]:
    rows = int(_safe_get(lambda: used_range.Rows.Count, 0) or 0)
    columns = int(_safe_get(lambda: used_range.Columns.Count, 0) or 0)
    cells: list[dict[str, Any]] = []
    merged_ranges: set[str] = set()

    for row_index in range(1, rows + 1):
        for column_index in range(1, columns + 1):
            cell = used_range.Cells(row_index, column_index)
            snapshot = _cell_snapshot(cell, str(sheet.Name))
            if snapshot is not None:
                cells.append(snapshot)
            merged_range = snapshot.get("merged_range") if snapshot else None
            if merged_range:
                merged_ranges.add(str(merged_range))

    return cells, sorted(merged_ranges)


def _collect_dimensions(sheet: Any, used_range: Any) -> dict[str, Any]:
    first_row = int(_safe_get(lambda: used_range.Row, 1) or 1)
    first_column = int(_safe_get(lambda: used_range.Column, 1) or 1)
    row_count = int(_safe_get(lambda: used_range.Rows.Count, 0) or 0)
    column_count = int(_safe_get(lambda: used_range.Columns.Count, 0) or 0)

    rows = []
    for index in range(first_row, first_row + row_count):
        row = sheet.Rows(index)
        rows.append(
            {
                "row": index,
                "height": _json_value(_safe_get(lambda r=row: r.RowHeight)),
                "hidden": bool(_safe_get(lambda r=row: r.Hidden, False)),
            }
        )

    columns = []
    for index in range(first_column, first_column + column_count):
        column = sheet.Columns(index)
        columns.append(
            {
                "column": index,
                "width": _json_value(_safe_get(lambda c=column: c.ColumnWidth)),
                "hidden": bool(_safe_get(lambda c=column: c.Hidden, False)),
            }
        )

    return {"rows": rows, "columns": columns}


def _collect_page_breaks(sheet: Any) -> dict[str, list[int]]:
    horizontal: list[int] = []
    vertical: list[int] = []

    h_breaks = _safe_get(lambda: sheet.HPageBreaks)
    if h_breaks is not None:
        count = int(_safe_get(lambda: h_breaks.Count, 0) or 0)
        for index in range(1, count + 1):
            row = _safe_get(lambda i=index: h_breaks.Item(i).Location.Row)
            if row is not None:
                horizontal.append(int(row))

    v_breaks = _safe_get(lambda: sheet.VPageBreaks)
    if v_breaks is not None:
        count = int(_safe_get(lambda: v_breaks.Count, 0) or 0)
        for index in range(1, count + 1):
            column = _safe_get(lambda i=index: v_breaks.Item(i).Location.Column)
            if column is not None:
                vertical.append(int(column))

    return {"horizontal_rows": horizontal, "vertical_columns": vertical}


def _collect_shapes(sheet: Any) -> list[dict[str, Any]]:
    shapes: list[dict[str, Any]] = []
    collection = _safe_get(lambda: sheet.Shapes)
    if collection is None:
        return shapes

    count = int(_safe_get(lambda: collection.Count, 0) or 0)
    for index in range(1, count + 1):
        shape = collection.Item(index)
        shapes.append(
            {
                "name": _json_value(_safe_get(lambda s=shape: s.Name)),
                "type": _json_value(_safe_get(lambda s=shape: s.Type)),
                "alternative_text": _json_value(_safe_get(lambda s=shape: s.AlternativeText)),
                "left": _json_value(_safe_get(lambda s=shape: s.Left)),
                "top": _json_value(_safe_get(lambda s=shape: s.Top)),
                "width": _json_value(_safe_get(lambda s=shape: s.Width)),
                "height": _json_value(_safe_get(lambda s=shape: s.Height)),
                "placement": _json_value(_safe_get(lambda s=shape: s.Placement)),
                "top_left_cell": _safe_get(lambda s=shape: str(s.TopLeftCell.Address(False, False))),
                "bottom_right_cell": _safe_get(lambda s=shape: str(s.BottomRightCell.Address(False, False))),
            }
        )
    return shapes


def _page_setup(sheet: Any) -> dict[str, Any]:
    setup = sheet.PageSetup
    return {
        "print_area": _json_value(_safe_get(lambda: setup.PrintArea)),
        "print_title_rows": _json_value(_safe_get(lambda: setup.PrintTitleRows)),
        "print_title_columns": _json_value(_safe_get(lambda: setup.PrintTitleColumns)),
        "orientation": _json_value(_safe_get(lambda: setup.Orientation)),
        "paper_size": _json_value(_safe_get(lambda: setup.PaperSize)),
        "zoom": _json_value(_safe_get(lambda: setup.Zoom)),
        "fit_to_pages_wide": _json_value(_safe_get(lambda: setup.FitToPagesWide)),
        "fit_to_pages_tall": _json_value(_safe_get(lambda: setup.FitToPagesTall)),
        "center_horizontally": bool(_safe_get(lambda: setup.CenterHorizontally, False)),
        "center_vertically": bool(_safe_get(lambda: setup.CenterVertically, False)),
        "margins": {
            "left": _json_value(_safe_get(lambda: setup.LeftMargin)),
            "right": _json_value(_safe_get(lambda: setup.RightMargin)),
            "top": _json_value(_safe_get(lambda: setup.TopMargin)),
            "bottom": _json_value(_safe_get(lambda: setup.BottomMargin)),
            "header": _json_value(_safe_get(lambda: setup.HeaderMargin)),
            "footer": _json_value(_safe_get(lambda: setup.FooterMargin)),
        },
        "headers": {
            "left": _json_value(_safe_get(lambda: setup.LeftHeader)),
            "center": _json_value(_safe_get(lambda: setup.CenterHeader)),
            "right": _json_value(_safe_get(lambda: setup.RightHeader)),
        },
        "footers": {
            "left": _json_value(_safe_get(lambda: setup.LeftFooter)),
            "center": _json_value(_safe_get(lambda: setup.CenterFooter)),
            "right": _json_value(_safe_get(lambda: setup.RightFooter)),
        },
    }


def _defined_names(workbook: Any) -> list[dict[str, Any]]:
    names: list[dict[str, Any]] = []
    collection = _safe_get(lambda: workbook.Names)
    if collection is None:
        return names
    count = int(_safe_get(lambda: collection.Count, 0) or 0)
    for index in range(1, count + 1):
        item = collection.Item(index)
        names.append(
            {
                "name": _json_value(_safe_get(lambda i=item: i.Name)),
                "refers_to": _json_value(_safe_get(lambda i=item: i.RefersTo)),
                "visible": bool(_safe_get(lambda i=item: i.Visible, True)),
            }
        )
    return names


def inspect_template(template_path: Path, format_id: str) -> dict[str, Any]:
    if os.name != "nt":
        raise RuntimeError("Excel COM inspection requires Windows with Microsoft Excel installed.")
    if not template_path.exists():
        raise FileNotFoundError(f"Template file not found: {template_path}")

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required. Install the project with: python -m pip install -e ."
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(
            str(template_path.resolve()),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
        )

        worksheets: list[dict[str, Any]] = []
        for sheet_index in range(1, int(workbook.Worksheets.Count) + 1):
            sheet = workbook.Worksheets(sheet_index)
            used_range = sheet.UsedRange
            cells, merged_ranges = _collect_cells(sheet, used_range)
            worksheets.append(
                {
                    "name": str(sheet.Name),
                    "index": sheet_index,
                    "visible": _json_value(_safe_get(lambda s=sheet: s.Visible)),
                    "used_range": _address(used_range),
                    "used_rows": int(_safe_get(lambda u=used_range: u.Rows.Count, 0) or 0),
                    "used_columns": int(_safe_get(lambda u=used_range: u.Columns.Count, 0) or 0),
                    "page_setup": _page_setup(sheet),
                    "page_breaks": _collect_page_breaks(sheet),
                    "merged_ranges": merged_ranges,
                    "dimensions": _collect_dimensions(sheet, used_range),
                    "shapes": _collect_shapes(sheet),
                    "cells": cells,
                }
            )

        return {
            "schema_version": 1,
            "format_id": format_id,
            "template": {
                "path": str(template_path.resolve()),
                "name": str(workbook.Name),
                "file_format": _json_value(_safe_get(lambda: workbook.FileFormat)),
                "read_only": bool(_safe_get(lambda: workbook.ReadOnly, True)),
                "worksheets_count": int(workbook.Worksheets.Count),
                "defined_names": _defined_names(workbook),
            },
            "worksheets": worksheets,
        }
    finally:
        if workbook is not None:
            _safe_get(lambda: workbook.Close(SaveChanges=False))
        if excel is not None:
            _safe_get(lambda: excel.Quit())
        pythoncom.CoUninitialize()


def _md(value: Any, max_length: int = 140) -> str:
    if value is None:
        return ""
    text = re.sub(r"\s+", " ", str(value)).replace("|", "\\|").strip()
    if len(text) > max_length:
        return text[: max_length - 1] + "…"
    return text


def inventory_markdown(inventory: dict[str, Any]) -> str:
    template = inventory["template"]
    lines = [
        f"# Template inventory — {inventory['format_id']}",
        "",
        f"- File: `{template['name']}`",
        f"- Path: `{template['path']}`",
        f"- Excel file format code: `{template['file_format']}`",
        f"- Worksheets: `{template['worksheets_count']}`",
        "",
    ]

    for sheet in inventory["worksheets"]:
        lines.extend(
            [
                f"## {sheet['name']}",
                "",
                f"- Used range: `{sheet['used_range']}`",
                f"- Print area: `{sheet['page_setup']['print_area']}`",
                f"- Horizontal page breaks: `{sheet['page_breaks']['horizontal_rows']}`",
                f"- Vertical page breaks: `{sheet['page_breaks']['vertical_columns']}`",
                f"- Merged ranges: `{len(sheet['merged_ranges'])}`",
                f"- Shapes/images: `{len(sheet['shapes'])}`",
                f"- Non-empty/formula cells: `{len(sheet['cells'])}`",
                "",
            ]
        )

        if sheet["shapes"]:
            lines.extend(["### Shapes and images", "", "| Name | Type | Anchor | Size |", "|---|---:|---|---|"])
            for shape in sheet["shapes"]:
                size = f"{_md(shape['width'])} × {_md(shape['height'])}"
                anchor = f"{_md(shape['top_left_cell'])} → {_md(shape['bottom_right_cell'])}"
                lines.append(f"| {_md(shape['name'])} | {_md(shape['type'])} | {anchor} | {size} |")
            lines.append("")

        if sheet["name"] in {"Información", "Resultados"}:
            lines.extend(
                [
                    "### Relevant cells",
                    "",
                    "| Cell | Value/display | Type | Formula | Merge | Probable role |",
                    "|---|---|---|---|---|---|",
                ]
            )
            for cell in sheet["cells"]:
                value = cell["display_text"] or cell["value"]
                lines.append(
                    "| "
                    f"{_md(cell['address'])} | {_md(value)} | {_md(cell['cell_type'])} | "
                    f"{_md(cell['formula'])} | {_md(cell['merged_range'])} | {_md(cell['probable_role'])} |"
                )
            lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def write_inventory(inventory: dict[str, Any], output_dir: Path) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    prefix = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(inventory["format_id"]))
    json_path = output_dir / f"{prefix}_inventory.json"
    markdown_path = output_dir / f"{prefix}_inventory.md"

    json_path.write_text(json.dumps(inventory, ensure_ascii=False, indent=2), encoding="utf-8")
    markdown_path.write_text(inventory_markdown(inventory), encoding="utf-8")
    return json_path, markdown_path
