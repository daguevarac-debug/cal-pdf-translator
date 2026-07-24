from __future__ import annotations

import json
import os
import re
from importlib import resources
from pathlib import Path
from typing import Any, Iterator

from cal_translator.formats.t50_04002 import FORMAT_ID


def _load_yaml(filename: str) -> dict[str, Any]:
    try:
        import yaml
    except ImportError as exc:
        raise RuntimeError(
            "PyYAML is required. Install the project with: python -m pip install -e ."
        ) from exc

    resource = resources.files(__package__).joinpath(filename)
    with resource.open("r", encoding="utf-8") as handle:
        payload = yaml.safe_load(handle)
    if not isinstance(payload, dict):
        raise RuntimeError(f"Invalid YAML document: {filename}")
    return payload


def _normalize_print_area(value: Any) -> str:
    text = str(value or "").strip()
    if "!" in text:
        text = text.rsplit("!", 1)[-1]
    text = text.replace("$", "").replace("'", "").replace(" ", "")
    return text.upper()


def _normalize_formula(value: Any) -> str:
    text = str(value or "").strip().replace("$", "")
    return re.sub(r"\s+", "", text).upper()


def _normalize_text(value: Any) -> str:
    text = str(value or "").replace("\u00a0", " ")
    return re.sub(r"\s+", " ", text).strip()


def _split_location(location: str) -> tuple[str, str]:
    try:
        sheet, address = location.rsplit("!", 1)
    except ValueError as exc:
        raise RuntimeError(f"Invalid contract location: {location}") from exc
    return sheet, address


def _mapped_locations(node: Any, path: tuple[str, ...] = ()) -> Iterator[tuple[str, str, str]]:
    if isinstance(node, dict):
        sheet = node.get("sheet")
        if isinstance(sheet, str):
            for key in ("cell", "range", "header_range", "dynamic_range"):
                address = node.get(key)
                if isinstance(address, str):
                    yield ".".join((*path, key)), sheet, address
        for key, value in node.items():
            yield from _mapped_locations(value, (*path, str(key)))
    elif isinstance(node, list):
        for index, value in enumerate(node):
            yield from _mapped_locations(value, (*path, str(index)))


def _add_check(
    report: dict[str, Any],
    *,
    check_type: str,
    target: str,
    expected: Any,
    actual: Any,
    passed: bool,
    message: str,
    severity: str = "error",
) -> None:
    check = {
        "type": check_type,
        "target": target,
        "expected": expected,
        "actual": actual,
        "passed": passed,
        "severity": severity,
        "message": message,
    }
    report["checks"].append(check)
    if passed:
        return
    issue = {
        "code": check_type,
        "target": target,
        "message": message,
        "expected": expected,
        "actual": actual,
    }
    report["warnings" if severity == "warning" else "errors"].append(issue)


def _worksheet_names(workbook: Any) -> list[str]:
    return [str(workbook.Worksheets(index).Name) for index in range(1, workbook.Worksheets.Count + 1)]


def _read_cell(workbook: Any, sheet_name: str, address: str) -> Any:
    return workbook.Worksheets(sheet_name).Range(address)


def _range_address(com_range: Any) -> str:
    address_member = com_range.Address
    if callable(address_member):
        return str(address_member(False, False))
    return str(address_member).replace("$", "")


def validate_template(template_path: Path, format_id: str = FORMAT_ID) -> dict[str, Any]:
    if format_id != FORMAT_ID:
        raise RuntimeError(f"Unsupported format: {format_id}. Expected {FORMAT_ID}.")
    if os.name != "nt":
        raise RuntimeError("Template validation requires Windows with Microsoft Excel installed.")
    if not template_path.exists():
        raise FileNotFoundError(f"Template file not found: {template_path}")

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required. Install the project with: python -m pip install -e ."
        ) from exc

    contract = _load_yaml("template_contract.yaml")
    cell_map = _load_yaml("cell_map.yaml")
    report: dict[str, Any] = {
        "valid": False,
        "format_id": FORMAT_ID,
        "template": str(template_path.resolve()),
        "errors": [],
        "warnings": [],
        "checks": [],
    }

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        workbook = excel.Workbooks.Open(
            str(template_path.resolve()),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
            IgnoreReadOnlyRecommended=True,
        )

        _add_check(
            report,
            check_type="workbook_read_only",
            target=str(template_path.name),
            expected=True,
            actual=bool(workbook.ReadOnly),
            passed=bool(workbook.ReadOnly),
            severity="warning",
            message="Workbook was opened in read-only mode." if bool(workbook.ReadOnly) else "Workbook did not report read-only mode.",
        )

        actual_sheets = _worksheet_names(workbook)
        actual_sheet_set = set(actual_sheets)
        for expected_sheet in contract.get("worksheets", []):
            present = expected_sheet in actual_sheet_set
            _add_check(
                report,
                check_type="required_worksheet",
                target=expected_sheet,
                expected="present",
                actual="present" if present else "missing",
                passed=present,
                message=(
                    f"Required worksheet '{expected_sheet}' is present."
                    if present
                    else f"Required worksheet '{expected_sheet}' is missing."
                ),
            )

        for sheet_name, expected_area in contract.get("print_areas", {}).items():
            if sheet_name not in actual_sheet_set:
                continue
            actual_area = workbook.Worksheets(sheet_name).PageSetup.PrintArea
            passed = _normalize_print_area(actual_area) == _normalize_print_area(expected_area)
            _add_check(
                report,
                check_type="print_area",
                target=sheet_name,
                expected=expected_area,
                actual=actual_area,
                passed=passed,
                message=(
                    f"Print area for '{sheet_name}' matches the controlled contract."
                    if passed
                    else f"Print area for '{sheet_name}' differs from the controlled contract."
                ),
            )

        for location, expected_formula in contract.get("required_formulas", {}).items():
            sheet_name, address = _split_location(location)
            if sheet_name not in actual_sheet_set:
                continue
            try:
                cell = _read_cell(workbook, sheet_name, address)
                actual_formula = cell.Formula if bool(cell.HasFormula) else None
                passed = _normalize_formula(actual_formula) == _normalize_formula(expected_formula)
            except Exception as exc:  # noqa: BLE001
                actual_formula = f"unresolvable: {exc}"
                passed = False
            _add_check(
                report,
                check_type="required_formula",
                target=location,
                expected=expected_formula,
                actual=actual_formula,
                passed=passed,
                message=(
                    f"Formula at {location} matches the controlled contract."
                    if passed
                    else f"Formula at {location} is missing or incompatible."
                ),
            )

        for location, expected_value in contract.get("required_values", {}).items():
            sheet_name, address = _split_location(location)
            if sheet_name not in actual_sheet_set:
                continue
            try:
                actual_value = _read_cell(workbook, sheet_name, address).Value2
                passed = _normalize_text(actual_value) == _normalize_text(expected_value)
            except Exception as exc:  # noqa: BLE001
                actual_value = f"unresolvable: {exc}"
                passed = False
            _add_check(
                report,
                check_type="required_anchor_value",
                target=location,
                expected=expected_value,
                actual=actual_value,
                passed=passed,
                message=(
                    f"Anchor text at {location} matches the controlled contract."
                    if passed
                    else f"Anchor text at {location} is missing or incompatible."
                ),
            )

        seen_locations: set[tuple[str, str]] = set()
        for logical_name, sheet_name, address in _mapped_locations(cell_map):
            key = (sheet_name, address)
            if key in seen_locations:
                continue
            seen_locations.add(key)
            if sheet_name not in actual_sheet_set:
                _add_check(
                    report,
                    check_type="mapped_location",
                    target=logical_name,
                    expected=f"{sheet_name}!{address}",
                    actual="worksheet missing",
                    passed=False,
                    message=f"Mapped location {logical_name} cannot be resolved because worksheet '{sheet_name}' is missing.",
                )
                continue
            try:
                resolved = _range_address(workbook.Worksheets(sheet_name).Range(address))
                passed = bool(resolved)
                actual = f"{sheet_name}!{resolved}" if resolved else None
            except Exception as exc:  # noqa: BLE001
                passed = False
                actual = f"unresolvable: {exc}"
            _add_check(
                report,
                check_type="mapped_location",
                target=logical_name,
                expected=f"{sheet_name}!{address}",
                actual=actual,
                passed=passed,
                message=(
                    f"Mapped location {logical_name} resolves successfully."
                    if passed
                    else f"Mapped location {logical_name} no longer resolves in the template."
                ),
            )

        critical_codes = {
            "required_worksheet",
            "print_area",
            "required_formula",
            "required_anchor_value",
            "mapped_location",
        }
        incompatible = any(error.get("code") in critical_codes for error in report["errors"])
        if incompatible:
            report["errors"].append(
                {
                    "code": "incompatible_template_revision",
                    "target": FORMAT_ID,
                    "message": "The workbook does not satisfy the controlled T50-04002 contract and may belong to an incompatible revision.",
                    "expected": "controlled T50-04002 template",
                    "actual": "contract mismatch",
                }
            )

        report["valid"] = not report["errors"]
        return report
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def _markdown_value(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("|", "\\|").replace("\n", " ")


def write_validation_report(report: dict[str, Any], output_dir: Path) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = f"{report['format_id']}_validation"
    json_path = output_dir / f"{stem}.json"
    markdown_path = output_dir / f"{stem}.md"

    json_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    lines = [
        f"# Template validation: {report['format_id']}",
        "",
        f"- Valid: `{str(report['valid']).lower()}`",
        f"- Template: `{report.get('template', '')}`",
        f"- Errors: {len(report['errors'])}",
        f"- Warnings: {len(report['warnings'])}",
        f"- Checks: {len(report['checks'])}",
        "",
        "## Checks",
        "",
        "| Status | Type | Target | Expected | Actual |",
        "|---|---|---|---|---|",
    ]
    for check in report["checks"]:
        status = "PASS" if check["passed"] else check["severity"].upper()
        lines.append(
            "| "
            + " | ".join(
                _markdown_value(value)
                for value in (
                    status,
                    check["type"],
                    check["target"],
                    check["expected"],
                    check["actual"],
                )
            )
            + " |"
        )

    if report["errors"]:
        lines.extend(["", "## Errors", ""])
        lines.extend(f"- `{item['code']}` {item['target']}: {item['message']}" for item in report["errors"])
    if report["warnings"]:
        lines.extend(["", "## Warnings", ""])
        lines.extend(f"- `{item['code']}` {item['target']}: {item['message']}" for item in report["warnings"])

    markdown_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return json_path, markdown_path
