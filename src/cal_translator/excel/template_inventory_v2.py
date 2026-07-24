from __future__ import annotations

import json
import logging
import os
import re
from pathlib import Path
from typing import Any, Iterable

from cal_translator.excel import template_inventory as legacy

LOGGER = logging.getLogger(__name__)

XL_CELL_TYPE_CONSTANTS = 2
XL_CELL_TYPE_FORMULAS = -4123


def _column_name(index: int) -> str:
    letters: list[str] = []
    value = index
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def _a1(row: int, column: int) -> str:
    return f"{_column_name(column)}{row}"


def _range_bounds(range_obj: Any) -> tuple[int, int, int, int]:
    first_row = int(legacy._safe_get(lambda: range_obj.Row, 1) or 1)
    first_column = int(legacy._safe_get(lambda: range_obj.Column, 1) or 1)
    row_count = int(legacy._safe_get(lambda: range_obj.Rows.Count, 0) or 0)
    column_count = int(legacy._safe_get(lambda: range_obj.Columns.Count, 0) or 0)
    last_row = first_row + max(0, row_count - 1)
    last_column = first_column + max(0, column_count - 1)
    return first_row, first_column, last_row, last_column


def _range_address(range_obj: Any) -> str:
    attempts = (
        lambda: range_obj.get_Address(False, False),
        lambda: range_obj.Address(False, False),
        lambda: range_obj.Address,
    )
    for getter in attempts:
        value = legacy._safe_get(getter)
        if value not in (None, ""):
            return str(value).replace("$", "")

    first_row, first_column, last_row, last_column = _range_bounds(range_obj)
    start = _a1(first_row, first_column)
    end = _a1(last_row, last_column)
    return start if start == end else f"{start}:{end}"


def _range_info(range_obj: Any) -> dict[str, Any]:
    first_row, first_column, last_row, last_column = _range_bounds(range_obj)
    return {
        "address": _range_address(range_obj),
        "first_row": first_row,
        "first_column": first_column,
        "last_row": last_row,
        "last_column": last_column,
        "rows": last_row - first_row + 1,
        "columns": last_column - first_column + 1,
    }


def _inside(row: int, column: int, bounds: tuple[int, int, int, int]) -> bool:
    first_row, first_column, last_row, last_column = bounds
    return first_row <= row <= last_row and first_column <= column <= last_column


def _print_range(sheet: Any, print_area: str | None, used_range: Any) -> Any:
    if print_area:
        resolved = legacy._safe_get(lambda: sheet.Range(print_area))
        if resolved is not None:
            return resolved
    return used_range


def _iter_range_cells(range_obj: Any) -> Iterable[Any]:
    areas = legacy._safe_get(lambda: range_obj.Areas)
    if areas is None:
        count = int(legacy._safe_get(lambda: range_obj.Cells.Count, 0) or 0)
        for index in range(1, count + 1):
            yield range_obj.Cells(index)
        return

    area_count = int(legacy._safe_get(lambda: areas.Count, 0) or 0)
    for area_index in range(1, area_count + 1):
        area = areas.Item(area_index)
        cell_count = int(legacy._safe_get(lambda a=area: a.Cells.Count, 0) or 0)
        for cell_index in range(1, cell_count + 1):
            yield area.Cells(cell_index)


def _special_cells(used_range: Any, cell_type: int) -> Iterable[Any]:
    selected = legacy._safe_get(lambda: used_range.SpecialCells(cell_type))
    if selected is None:
        return ()
    return _iter_range_cells(selected)


def _cell_snapshot(cell: Any, sheet_name: str, scope: str) -> dict[str, Any] | None:
    snapshot = legacy._cell_snapshot(cell, sheet_name)
    if snapshot is None:
        return None

    row = int(legacy._safe_get(lambda: cell.Row, 0) or 0)
    column = int(legacy._safe_get(lambda: cell.Column, 0) or 0)
    snapshot["address"] = _a1(row, column)
    snapshot["row"] = row
    snapshot["column"] = column
    snapshot["scope"] = scope

    if bool(legacy._safe_get(lambda: cell.MergeCells, False)):
        merge_area = legacy._safe_get(lambda: cell.MergeArea)
        snapshot["merged_range"] = _range_address(merge_area) if merge_area is not None else None

    return snapshot


def _collect_cells(
    sheet: Any,
    used_range: Any,
    print_range: Any,
) -> tuple[list[dict[str, Any]], list[str]]:
    print_bounds = _range_bounds(print_range)
    seen: set[tuple[int, int]] = set()
    cells: list[dict[str, Any]] = []
    merged_ranges: set[str] = set()

    candidates = list(_iter_range_cells(print_range))
    candidates.extend(_special_cells(used_range, XL_CELL_TYPE_CONSTANTS))
    candidates.extend(_special_cells(used_range, XL_CELL_TYPE_FORMULAS))

    for cell in candidates:
        row = int(legacy._safe_get(lambda c=cell: c.Row, 0) or 0)
        column = int(legacy._safe_get(lambda c=cell: c.Column, 0) or 0)
        key = (row, column)
        if key in seen or row <= 0 or column <= 0:
            continue
        seen.add(key)

        scope = "print_area" if _inside(row, column, print_bounds) else "auxiliary"
        snapshot = _cell_snapshot(cell, str(sheet.Name), scope)
        if snapshot is None:
            continue
        cells.append(snapshot)
        merged_range = snapshot.get("merged_range")
        if merged_range:
            merged_ranges.add(str(merged_range))

    cells.sort(key=lambda item: (int(item["row"]), int(item["column"])))
    return cells, sorted(merged_ranges)


def _collect_shapes(sheet: Any) -> list[dict[str, Any]]:
    shapes = legacy._collect_shapes(sheet)
    collection = legacy._safe_get(lambda: sheet.Shapes)
    if collection is None:
        return shapes

    count = int(legacy._safe_get(lambda: collection.Count, 0) or 0)
    for index, snapshot in enumerate(shapes, start=1):
        if index > count:
            break
        shape = collection.Item(index)
        top_left = legacy._safe_get(lambda s=shape: s.TopLeftCell)
        bottom_right = legacy._safe_get(lambda s=shape: s.BottomRightCell)
        snapshot["top_left_cell"] = _range_address(top_left) if top_left is not None else None
        snapshot["bottom_right_cell"] = _range_address(bottom_right) if bottom_right is not None else None
    return shapes


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
        sheet_count = int(workbook.Worksheets.Count)
        for sheet_index in range(1, sheet_count + 1):
            sheet = workbook.Worksheets(sheet_index)
            LOGGER.info("Inspecting worksheet %s/%s: %s", sheet_index, sheet_count, sheet.Name)
            used_range = sheet.UsedRange
            page_setup = legacy._page_setup(sheet)
            print_range = _print_range(sheet, page_setup.get("print_area"), used_range)
            cells, merged_ranges = _collect_cells(sheet, used_range, print_range)
            print_cells = sum(1 for cell in cells if cell["scope"] == "print_area")
            auxiliary_cells = len(cells) - print_cells

            worksheets.append(
                {
                    "name": str(sheet.Name),
                    "index": sheet_index,
                    "visible": legacy._json_value(legacy._safe_get(lambda s=sheet: s.Visible)),
                    "used_range": _range_info(used_range),
                    "print_area_range": _range_info(print_range),
                    "page_setup": page_setup,
                    "page_breaks": legacy._collect_page_breaks(sheet),
                    "merged_ranges": merged_ranges,
                    "print_dimensions": legacy._collect_dimensions(sheet, print_range),
                    "shapes": _collect_shapes(sheet),
                    "cell_counts": {
                        "total": len(cells),
                        "print_area": print_cells,
                        "auxiliary": auxiliary_cells,
                    },
                    "cells": cells,
                }
            )

        return {
            "schema_version": 2,
            "format_id": format_id,
            "template": {
                "path": str(template_path.resolve()),
                "name": str(workbook.Name),
                "file_format": legacy._json_value(legacy._safe_get(lambda: workbook.FileFormat)),
                "read_only": bool(legacy._safe_get(lambda: workbook.ReadOnly, True)),
                "worksheets_count": sheet_count,
                "defined_names": legacy._defined_names(workbook),
            },
            "worksheets": worksheets,
        }
    finally:
        if workbook is not None:
            legacy._safe_get(lambda: workbook.Close(SaveChanges=False))
        if excel is not None:
            legacy._safe_get(lambda: excel.Quit())
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
        f"- Inventory schema: `{inventory['schema_version']}`",
        "",
    ]

    for sheet in inventory["worksheets"]:
        used = sheet["used_range"]
        printed = sheet["print_area_range"]
        counts = sheet["cell_counts"]
        lines.extend(
            [
                f"## {sheet['name']}",
                "",
                f"- Used range: `{used['address']}` ({used['rows']} × {used['columns']})",
                f"- Print area range: `{printed['address']}` ({printed['rows']} × {printed['columns']})",
                f"- Horizontal page breaks: `{sheet['page_breaks']['horizontal_rows']}`",
                f"- Vertical page breaks: `{sheet['page_breaks']['vertical_columns']}`",
                f"- Merged ranges: `{len(sheet['merged_ranges'])}`",
                f"- Shapes/images: `{len(sheet['shapes'])}`",
                f"- Relevant cells: `{counts['total']}` (print area: `{counts['print_area']}`, auxiliary: `{counts['auxiliary']}`)",
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
                    "| Cell | Scope | Value/display | Type | Formula | Merge | Probable role |",
                    "|---|---|---|---|---|---|---|",
                ]
            )
            for cell in sheet["cells"]:
                value = cell["display_text"] or cell["value"]
                lines.append(
                    "| "
                    f"{_md(cell['address'])} | {_md(cell['scope'])} | {_md(value)} | {_md(cell['cell_type'])} | "
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
