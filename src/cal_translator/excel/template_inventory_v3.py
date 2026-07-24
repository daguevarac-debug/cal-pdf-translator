from __future__ import annotations

import json
import logging
import os
import re
from pathlib import Path
from typing import Any, Iterable

from cal_translator.excel import template_inventory as legacy
from cal_translator.excel import template_inventory_v2 as v2

LOGGER = logging.getLogger(__name__)

XL_CELL_TYPE_CONSTANTS = 2
XL_CELL_TYPE_FORMULAS = -4123
XL_LINE_STYLE_NONE = -4142

_RESULT_MATRIX_SHEET = "Resultados"
_RESULT_MATRIX_ADDRESS = "B7:L24"

_A1_REFERENCE_RE = re.compile(
    r"(?:(?:'(?P<quoted_sheet>[^']+)'|(?P<plain_sheet>[A-Za-z_][A-Za-z0-9_ ]*))!)?"
    r"\$?(?P<column>[A-Z]{1,3})\$?(?P<row>\d+)",
    re.IGNORECASE,
)


def _inside(row: int, column: int, bounds: tuple[int, int, int, int]) -> bool:
    first_row, first_column, last_row, last_column = bounds
    return first_row <= row <= last_row and first_column <= column <= last_column


def _scope(row: int, column: int, print_bounds: tuple[int, int, int, int]) -> str:
    return "print_area" if _inside(row, column, print_bounds) else "auxiliary"


def _special_cells(used_range: Any, cell_type: int) -> Iterable[Any]:
    selected = legacy._safe_get(lambda: used_range.SpecialCells(cell_type))
    if selected is None:
        return ()
    return v2._iter_range_cells(selected)


def _meaningful_borders(cell: Any) -> dict[str, Any]:
    borders = legacy._border_snapshot(cell)
    return {
        name: snapshot
        for name, snapshot in borders.items()
        if snapshot.get("line_style") not in (None, 0, XL_LINE_STYLE_NONE)
    }


def _merge_metadata(cell: Any) -> tuple[str | None, str | None, bool]:
    if not bool(legacy._safe_get(lambda: cell.MergeCells, False)):
        return None, None, False

    merge_area = legacy._safe_get(lambda: cell.MergeArea)
    if merge_area is None:
        return None, None, False

    merged_range = v2._range_address(merge_area)
    first_row = int(legacy._safe_get(lambda: merge_area.Row, 0) or 0)
    first_column = int(legacy._safe_get(lambda: merge_area.Column, 0) or 0)
    anchor = v2._a1(first_row, first_column)
    row = int(legacy._safe_get(lambda: cell.Row, 0) or 0)
    column = int(legacy._safe_get(lambda: cell.Column, 0) or 0)
    return merged_range, anchor, row == first_row and column == first_column


def _format_snapshot(cell: Any) -> dict[str, Any]:
    return {
        "number_format": legacy._json_value(legacy._safe_get(lambda: cell.NumberFormatLocal)),
        "style": legacy._json_value(legacy._safe_get(lambda: cell.Style)),
        "alignment": {
            "horizontal": legacy._json_value(legacy._safe_get(lambda: cell.HorizontalAlignment)),
            "vertical": legacy._json_value(legacy._safe_get(lambda: cell.VerticalAlignment)),
            "wrap_text": bool(legacy._safe_get(lambda: cell.WrapText, False)),
            "orientation": legacy._json_value(legacy._safe_get(lambda: cell.Orientation)),
        },
        "font": {
            "name": legacy._json_value(legacy._safe_get(lambda: cell.Font.Name)),
            "size": legacy._json_value(legacy._safe_get(lambda: cell.Font.Size)),
            "bold": bool(legacy._safe_get(lambda: cell.Font.Bold, False)),
            "italic": bool(legacy._safe_get(lambda: cell.Font.Italic, False)),
            "color": legacy._json_value(legacy._safe_get(lambda: cell.Font.Color)),
        },
        "fill_color": legacy._json_value(legacy._safe_get(lambda: cell.Interior.Color)),
        "borders": _meaningful_borders(cell),
    }


def _has_meaningful_format(formatting: dict[str, Any], merged_range: str | None) -> bool:
    style = str(formatting.get("style") or "").strip()
    number_format = str(formatting.get("number_format") or "").strip()
    alignment = formatting.get("alignment") or {}
    font = formatting.get("font") or {}
    fill_color = formatting.get("fill_color")

    return any(
        (
            merged_range is not None,
            style not in {"", "Normal"},
            number_format not in {"", "General"},
            bool(formatting.get("borders")),
            fill_color not in (None, 0, 16777215, 16777215.0),
            bool(alignment.get("wrap_text")),
            bool(font.get("bold")),
            bool(font.get("italic")),
        )
    )


def _cell_snapshot(
    cell: Any,
    sheet_name: str,
    scope: str,
    *,
    force: bool = False,
) -> dict[str, Any] | None:
    row = int(legacy._safe_get(lambda: cell.Row, 0) or 0)
    column = int(legacy._safe_get(lambda: cell.Column, 0) or 0)
    if row <= 0 or column <= 0:
        return None

    merged_range, merge_anchor, is_merge_anchor = _merge_metadata(cell)
    if merged_range is not None and not is_merge_anchor:
        return None

    has_formula = bool(legacy._safe_get(lambda: cell.HasFormula, False))
    value = legacy._safe_get(lambda: cell.Value2)
    text = str(legacy._safe_get(lambda: cell.Text, "") or "")
    comment = legacy._safe_get(lambda: cell.Comment)
    comment_text = legacy._safe_get(lambda: comment.Text(), None) if comment is not None else None
    formatting = _format_snapshot(cell)
    is_blank = value in (None, "") and not has_formula and not comment_text

    if is_blank and not force and not _has_meaningful_format(formatting, merged_range):
        return None

    formula = legacy._safe_get(lambda: cell.Formula, None) if has_formula else None
    snapshot = {
        "address": v2._a1(row, column),
        "row": row,
        "column": column,
        "scope": scope,
        "value": legacy._json_value(value),
        "display_text": text,
        "cell_type": legacy._cell_type(value, has_formula),
        "is_blank": is_blank,
        "formula": legacy._json_value(formula),
        "merged_range": merged_range,
        "merge_anchor": merge_anchor,
        "is_merge_anchor": is_merge_anchor,
        "probable_role": legacy._probable_role(sheet_name, text, value, has_formula),
        "comment": legacy._json_value(comment_text),
    }
    snapshot.update(formatting)
    return snapshot


def _collect_cells(
    sheet: Any,
    used_range: Any,
    print_range: Any,
) -> tuple[list[dict[str, Any]], list[str], set[tuple[int, int]]]:
    print_bounds = v2._range_bounds(print_range)
    seen: set[tuple[int, int]] = set()
    cells: list[dict[str, Any]] = []
    merged_ranges: set[str] = set()

    candidates = list(v2._iter_range_cells(print_range))
    candidates.extend(_special_cells(used_range, XL_CELL_TYPE_CONSTANTS))
    candidates.extend(_special_cells(used_range, XL_CELL_TYPE_FORMULAS))

    for cell in candidates:
        row = int(legacy._safe_get(lambda c=cell: c.Row, 0) or 0)
        column = int(legacy._safe_get(lambda c=cell: c.Column, 0) or 0)
        key = (row, column)
        if key in seen or row <= 0 or column <= 0:
            continue
        seen.add(key)

        snapshot = _cell_snapshot(
            cell,
            str(sheet.Name),
            _scope(row, column, print_bounds),
        )
        if snapshot is None:
            continue
        cells.append(snapshot)
        if snapshot.get("merged_range"):
            merged_ranges.add(str(snapshot["merged_range"]))

    cells.sort(key=lambda item: (int(item["row"]), int(item["column"])))
    return cells, sorted(merged_ranges), seen


def _formula_references(formula: str, current_sheet: str) -> list[dict[str, str]]:
    references: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()

    for match in _A1_REFERENCE_RE.finditer(formula.upper()):
        sheet_name = (match.group("quoted_sheet") or match.group("plain_sheet") or current_sheet).strip()
        address = f"{match.group('column').upper()}{match.group('row')}"
        key = (sheet_name.casefold(), address)
        if key in seen:
            continue
        seen.add(key)
        references.append({"sheet": sheet_name, "address": address})

    return references


def _collect_formula_dependencies(
    sheet: Any,
    cells: list[dict[str, Any]],
    print_bounds: tuple[int, int, int, int],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    sheet_name = str(sheet.Name)
    dependencies: list[dict[str, Any]] = []
    referenced_local: dict[str, list[str]] = {}

    for cell in cells:
        formula = cell.get("formula")
        if not formula:
            continue
        references = _formula_references(str(formula), sheet_name)
        dependencies.append(
            {
                "formula_cell": cell["address"],
                "formula": formula,
                "references": references,
            }
        )
        for reference in references:
            if reference["sheet"].casefold() == sheet_name.casefold():
                referenced_local.setdefault(reference["address"], []).append(str(cell["address"]))

    existing = {str(cell["address"]): cell for cell in cells}
    dependency_cells: list[dict[str, Any]] = []
    for address, consumers in sorted(referenced_local.items()):
        target = existing.get(address)
        if target is None:
            com_cell = legacy._safe_get(lambda a=address: sheet.Range(a))
            if com_cell is None:
                continue
            row = int(legacy._safe_get(lambda c=com_cell: c.Row, 0) or 0)
            column = int(legacy._safe_get(lambda c=com_cell: c.Column, 0) or 0)
            target = _cell_snapshot(
                com_cell,
                sheet_name,
                _scope(row, column, print_bounds),
                force=True,
            )
            if target is None:
                continue
            cells.append(target)
            existing[address] = target

        dependency_cells.append(
            {
                "address": address,
                "scope": target["scope"],
                "is_blank": bool(target["is_blank"]),
                "display_text": target["display_text"],
                "value": target["value"],
                "number_format": target["number_format"],
                "consumed_by": sorted(set(consumers)),
            }
        )

    cells.sort(key=lambda item: (int(item["row"]), int(item["column"])))
    return dependencies, dependency_cells


def _result_matrix(sheet: Any, print_bounds: tuple[int, int, int, int]) -> dict[str, Any] | None:
    if str(sheet.Name).casefold() != _RESULT_MATRIX_SHEET.casefold():
        return None

    matrix_range = sheet.Range(_RESULT_MATRIX_ADDRESS)
    first_row, first_column, last_row, last_column = v2._range_bounds(matrix_range)
    cells: list[dict[str, Any]] = []

    for com_cell in v2._iter_range_cells(matrix_range):
        snapshot = _cell_snapshot(
            com_cell,
            str(sheet.Name),
            _scope(
                int(legacy._safe_get(lambda c=com_cell: c.Row, 0) or 0),
                int(legacy._safe_get(lambda c=com_cell: c.Column, 0) or 0),
                print_bounds,
            ),
            force=True,
        )
        if snapshot is None:
            continue
        snapshot["matrix_role"] = "header" if int(snapshot["row"]) == first_row else "data_slot"
        snapshot["writable_candidate"] = (
            int(snapshot["row"]) > first_row and not bool(snapshot.get("formula"))
        )
        cells.append(snapshot)

    cells.sort(key=lambda item: (int(item["row"]), int(item["column"])))
    return {
        "address": _RESULT_MATRIX_ADDRESS,
        "first_row": first_row,
        "last_row": last_row,
        "first_column": first_column,
        "last_column": last_column,
        "data_rows": max(0, last_row - first_row),
        "columns": last_column - first_column + 1,
        "cells": cells,
    }


def _writable_candidates(
    sheet_name: str,
    cells: list[dict[str, Any]],
    dependency_cells: list[dict[str, Any]],
    result_matrix: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    dependency_by_address = {item["address"]: item for item in dependency_cells}
    matrix_addresses = {
        str(cell["address"])
        for cell in (result_matrix or {}).get("cells", [])
        if cell.get("writable_candidate")
    }
    candidates: list[dict[str, Any]] = []

    for cell in cells:
        if cell.get("formula"):
            continue

        reasons: list[str] = []
        address = str(cell["address"])
        if cell.get("probable_role") == "variable_placeholder":
            reasons.append("variable_placeholder")
        if address in dependency_by_address and bool(cell.get("is_blank")):
            reasons.append("formula_input")
        if address in matrix_addresses:
            reasons.append("result_matrix_slot")
        if (
            bool(cell.get("is_blank"))
            and cell.get("scope") == "print_area"
            and _has_meaningful_format(cell, cell.get("merged_range"))
        ):
            reasons.append("formatted_blank")

        if not reasons:
            continue

        candidates.append(
            {
                "address": address,
                "scope": cell["scope"],
                "merged_range": cell.get("merged_range"),
                "is_merge_anchor": bool(cell.get("is_merge_anchor")),
                "current_display": cell.get("display_text"),
                "current_value": cell.get("value"),
                "number_format": cell.get("number_format"),
                "reasons": sorted(set(reasons)),
            }
        )

    candidates.sort(key=lambda item: v2._range_bounds(sheet_name and _FakeRange(item["address"])))
    return candidates


class _FakeRange:
    """Small adapter used only to sort A1 addresses without Excel COM."""

    def __init__(self, address: str) -> None:
        match = re.fullmatch(r"([A-Z]{1,3})(\d+)", address)
        if match is None:
            self.Row = 0
            self.Column = 0
            self.Rows = _FakeCount(1)
            self.Columns = _FakeCount(1)
            return
        self.Row = int(match.group(2))
        self.Column = _column_index(match.group(1))
        self.Rows = _FakeCount(1)
        self.Columns = _FakeCount(1)

    @property
    def Address(self) -> str:
        return v2._a1(self.Row, self.Column)


class _FakeCount:
    def __init__(self, count: int) -> None:
        self.Count = count


def _column_index(name: str) -> int:
    value = 0
    for character in name:
        value = value * 26 + ord(character) - 64
    return value


def _candidate_sort_key(item: dict[str, Any]) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]{1,3})(\d+)", str(item["address"]))
    if match is None:
        return (0, 0)
    return (int(match.group(2)), _column_index(match.group(1)))


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
            print_range = v2._print_range(sheet, page_setup.get("print_area"), used_range)
            print_bounds = v2._range_bounds(print_range)

            cells, merged_ranges, _ = _collect_cells(sheet, used_range, print_range)
            formula_dependencies, dependency_cells = _collect_formula_dependencies(
                sheet,
                cells,
                print_bounds,
            )
            result_matrix = _result_matrix(sheet, print_bounds)
            writable_candidates = _writable_candidates(
                str(sheet.Name),
                cells,
                dependency_cells,
                result_matrix,
            )
            writable_candidates.sort(key=_candidate_sort_key)

            print_cells = sum(1 for cell in cells if cell["scope"] == "print_area")
            auxiliary_cells = len(cells) - print_cells
            worksheet_inventory = {
                "name": str(sheet.Name),
                "index": sheet_index,
                "visible": legacy._json_value(legacy._safe_get(lambda s=sheet: s.Visible)),
                "used_range": v2._range_info(used_range),
                "print_area_range": v2._range_info(print_range),
                "page_setup": page_setup,
                "page_breaks": legacy._collect_page_breaks(sheet),
                "merged_ranges": merged_ranges,
                "print_dimensions": legacy._collect_dimensions(sheet, print_range),
                "shapes": v2._collect_shapes(sheet),
                "cell_counts": {
                    "total": len(cells),
                    "print_area": print_cells,
                    "auxiliary": auxiliary_cells,
                    "writable_candidates": len(writable_candidates),
                    "formula_dependencies": len(formula_dependencies),
                },
                "formula_dependencies": formula_dependencies,
                "dependency_cells": dependency_cells,
                "writable_candidates": writable_candidates,
                "cells": cells,
            }
            if result_matrix is not None:
                worksheet_inventory["result_entry_matrix"] = result_matrix
            worksheets.append(worksheet_inventory)

        return {
            "schema_version": 3,
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
                f"- Writable candidates: `{counts['writable_candidates']}`",
                f"- Formula cells with direct references: `{counts['formula_dependencies']}`",
                "",
            ]
        )

        if sheet["formula_dependencies"]:
            lines.extend(
                [
                    "### Formula dependencies",
                    "",
                    "| Formula cell | Formula | Direct references |",
                    "|---|---|---|",
                ]
            )
            for dependency in sheet["formula_dependencies"]:
                references = ", ".join(
                    f"{item['sheet']}!{item['address']}" for item in dependency["references"]
                )
                lines.append(
                    f"| {_md(dependency['formula_cell'])} | {_md(dependency['formula'])} | {_md(references)} |"
                )
            lines.append("")

        if sheet["writable_candidates"]:
            lines.extend(
                [
                    "### Writable candidates",
                    "",
                    "| Cell | Scope | Merge | Current value | Number format | Reasons |",
                    "|---|---|---|---|---|---|",
                ]
            )
            for candidate in sheet["writable_candidates"]:
                lines.append(
                    "| "
                    f"{_md(candidate['address'])} | {_md(candidate['scope'])} | {_md(candidate['merged_range'])} | "
                    f"{_md(candidate['current_display'] or candidate['current_value'])} | "
                    f"{_md(candidate['number_format'])} | {_md(', '.join(candidate['reasons']))} |"
                )
            lines.append("")

        matrix = sheet.get("result_entry_matrix")
        if matrix is not None:
            lines.extend(
                [
                    "### Result entry matrix",
                    "",
                    f"- Range: `{matrix['address']}`",
                    f"- Header rows: `1`; data rows available: `{matrix['data_rows']}`",
                    f"- Columns: `{matrix['columns']}`",
                    "",
                    "| Cell | Role | Value/display | Formula | Number format | Merge | Writable candidate |",
                    "|---|---|---|---|---|---|---|",
                ]
            )
            for cell in matrix["cells"]:
                lines.append(
                    "| "
                    f"{_md(cell['address'])} | {_md(cell['matrix_role'])} | "
                    f"{_md(cell['display_text'] or cell['value'])} | {_md(cell['formula'])} | "
                    f"{_md(cell['number_format'])} | {_md(cell['merged_range'])} | "
                    f"{_md(cell['writable_candidate'])} |"
                )
            lines.append("")

        if sheet["name"] in {"Información", "Resultados"}:
            lines.extend(
                [
                    "### Relevant cells",
                    "",
                    "| Cell | Scope | Blank | Merge anchor | Value/display | Type | Formula | Merge | Probable role |",
                    "|---|---|---|---|---|---|---|---|---|",
                ]
            )
            for cell in sheet["cells"]:
                value = cell["display_text"] or cell["value"]
                lines.append(
                    "| "
                    f"{_md(cell['address'])} | {_md(cell['scope'])} | {_md(cell['is_blank'])} | "
                    f"{_md(cell['is_merge_anchor'])} | {_md(value)} | {_md(cell['cell_type'])} | "
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
