from __future__ import annotations

import json
import re
import shutil
from dataclasses import asdict, dataclass
from datetime import datetime
from importlib import resources
from pathlib import Path
from typing import Any

from cal_translator.excel.com_backend import open_excel_workbook
from cal_translator.formats.t50_04002 import FORMAT_ID
from cal_translator.formats.t50_04002.validate_template import validate_template


@dataclass(slots=True)
class WriteOperation:
    logical_name: str
    sheet: str
    address: str
    value: Any


def _load_cell_map() -> dict[str, Any]:
    try:
        import yaml
    except ImportError as exc:
        raise RuntimeError(
            "PyYAML is required. Install the project with: python -m pip install -e ."
        ) from exc

    resource = resources.files(__package__).joinpath("cell_map.yaml")
    with resource.open("r", encoding="utf-8") as handle:
        payload = yaml.safe_load(handle)
    if not isinstance(payload, dict):
        raise RuntimeError("Invalid T50-04002 cell map.")
    return payload


def _node(cell_map: dict[str, Any], *path: str) -> dict[str, Any]:
    current: Any = cell_map
    for key in path:
        if not isinstance(current, dict) or key not in current:
            raise RuntimeError(f"Missing cell-map path: {'.'.join(path)}")
        current = current[key]
    if not isinstance(current, dict):
        raise RuntimeError(f"Invalid cell-map node: {'.'.join(path)}")
    return current


def _operation(
    cell_map: dict[str, Any],
    logical_name: str,
    path: tuple[str, ...],
    value: Any,
) -> WriteOperation | None:
    if value is None or value == "":
        return None
    target = _node(cell_map, *path)
    if not target.get("writable", False):
        raise RuntimeError(f"Cell-map target is not writable: {logical_name}")
    return WriteOperation(
        logical_name=logical_name,
        sheet=str(target["sheet"]),
        address=str(target["cell"]),
        value=value,
    )


def _number(value: Any) -> float | None:
    if value is None:
        return None
    match = re.search(r"[-+]?\d+(?:[.,]\d+)?", str(value))
    if not match:
        return None
    return float(match.group(0).replace(",", "."))


def _date(value: Any) -> datetime | None:
    if value is None or value == "":
        return None
    try:
        return datetime.strptime(str(value), "%Y-%m-%d")
    except ValueError as exc:
        raise RuntimeError(f"Invalid ISO date in extracted data: {value}") from exc


def build_scalar_write_plan(payload: dict[str, Any]) -> list[WriteOperation]:
    if payload.get("format_id") != FORMAT_ID:
        raise RuntimeError(
            f"Unsupported extracted-data format: {payload.get('format_id')}. Expected {FORMAT_ID}."
        )

    cell_map = _load_cell_map()
    equipment = payload.get("equipment") or {}
    customer = payload.get("customer") or {}
    environment = payload.get("environmental_conditions") or {}

    candidates: list[WriteOperation | None] = [
        _operation(cell_map, "certificate_number", ("certificate", "number_input"), payload.get("certificate_number")),
        _operation(cell_map, "internal_code", ("internal_code",), payload.get("internal_code")),
        _operation(cell_map, "equipment.description", ("equipment", "description"), equipment.get("description")),
        _operation(cell_map, "equipment.manufacturer", ("equipment", "manufacturer"), equipment.get("manufacturer")),
        _operation(cell_map, "equipment.serial_number", ("equipment", "serial_number"), equipment.get("serial_number")),
        _operation(cell_map, "equipment.model", ("equipment", "model"), equipment.get("model")),
        _operation(
            cell_map,
            "equipment.customer_identification",
            ("equipment", "customer_identification"),
            equipment.get("customer_identification"),
        ),
        _operation(cell_map, "customer.name", ("customer", "name"), customer.get("name")),
        _operation(cell_map, "customer.address", ("customer", "address"), customer.get("address")),
        _operation(cell_map, "customer.city", ("customer", "city"), customer.get("city")),
        _operation(cell_map, "customer.order_number", ("customer", "order_number"), customer.get("order_number")),
        _operation(cell_map, "dates.received", ("dates", "received"), _date(payload.get("reception_date"))),
        _operation(cell_map, "dates.calibration", ("dates", "calibration"), _date(payload.get("calibration_date"))),
        _operation(cell_map, "dates.issue", ("dates", "issue"), _date(payload.get("issue_date"))),
        _operation(
            cell_map,
            "environment.maximum_temperature.measured",
            ("environment", "maximum_temperature", "measured"),
            _number(environment.get("maximum_temperature")),
        ),
        _operation(
            cell_map,
            "environment.maximum_temperature.correction",
            ("environment", "maximum_temperature", "correction"),
            0.0,
        ),
        _operation(
            cell_map,
            "environment.minimum_temperature.measured",
            ("environment", "minimum_temperature", "measured"),
            _number(environment.get("minimum_temperature")),
        ),
        _operation(
            cell_map,
            "environment.minimum_temperature.correction",
            ("environment", "minimum_temperature", "correction"),
            0.0,
        ),
        _operation(
            cell_map,
            "environment.maximum_humidity.measured",
            ("environment", "maximum_humidity", "measured"),
            _number(environment.get("maximum_relative_humidity")),
        ),
        _operation(
            cell_map,
            "environment.maximum_humidity.correction",
            ("environment", "maximum_humidity", "correction"),
            0.0,
        ),
        _operation(
            cell_map,
            "environment.minimum_humidity.measured",
            ("environment", "minimum_humidity", "measured"),
            _number(environment.get("minimum_relative_humidity")),
        ),
        _operation(
            cell_map,
            "environment.minimum_humidity.correction",
            ("environment", "minimum_humidity", "correction"),
            0.0,
        ),
    ]
    return [operation for operation in candidates if operation is not None]


def _write_value(cell: Any, value: Any) -> None:
    if isinstance(value, datetime):
        cell.Value = value
    else:
        cell.Value2 = value


def build_workbook_prototype(
    template_path: Path,
    data_path: Path,
    output_path: Path,
    *,
    format_id: str = FORMAT_ID,
    overwrite: bool = False,
) -> tuple[dict[str, Any], Path]:
    if format_id != FORMAT_ID:
        raise RuntimeError(f"Unsupported workbook format: {format_id}. Expected {FORMAT_ID}.")
    if not data_path.exists():
        raise FileNotFoundError(f"Extracted JSON file not found: {data_path}")

    payload = json.loads(data_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise RuntimeError("Extracted certificate JSON must contain an object.")
    warnings = payload.get("extraction_warnings") or []
    if warnings:
        raise RuntimeError(
            "Workbook generation requires extraction_warnings to be empty. "
            f"Found {len(warnings)} warning(s)."
        )

    preflight = validate_template(template_path, format_id)
    if not preflight["valid"]:
        raise RuntimeError(
            f"Template validation failed with {len(preflight['errors'])} error(s)."
        )

    output_path = output_path.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists() and not overwrite:
        raise RuntimeError(f"Output workbook already exists: {output_path}")
    if output_path.suffix.lower() != ".xlsx":
        raise RuntimeError("Prototype workbook output must use the .xlsx extension.")

    shutil.copy2(template_path.resolve(), output_path)
    plan = build_scalar_write_plan(payload)

    try:
        with open_excel_workbook(output_path, read_only=False) as (excel, workbook):
            if bool(workbook.ReadOnly):
                raise RuntimeError("Copied workbook unexpectedly opened as read-only.")
            for operation in plan:
                cell = workbook.Worksheets(operation.sheet).Range(operation.address)
                _write_value(cell, operation.value)
            excel.CalculateFull()
            workbook.Save()
    except Exception:
        output_path.unlink(missing_ok=True)
        raise

    postflight = validate_template(output_path, format_id)
    report: dict[str, Any] = {
        "status": "prototype",
        "format_id": format_id,
        "source_template": str(template_path.resolve()),
        "source_data": str(data_path.resolve()),
        "output_workbook": str(output_path),
        "fields_written": len(plan),
        "write_operations": [asdict(operation) for operation in plan],
        "skipped_sections": {
            "traceability": {
                "rows": len(payload.get("traceability") or []),
                "reason": "repeatable row expansion remains pending validation",
            },
            "results": {
                "rows": len(payload.get("results") or []),
                "reason": "72-row result expansion remains pending validation",
            },
            "controlled_translation": {
                "reason": "English controlled-text replacement is not part of this prototype",
            },
        },
        "postflight_valid": bool(postflight["valid"]),
        "postflight_errors": postflight["errors"],
        "postflight_warnings": postflight["warnings"],
    }

    report_path = output_path.with_name(f"{output_path.stem}_build_report.json")
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    return report, report_path
