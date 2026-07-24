from __future__ import annotations

import json
import re
import shutil
from dataclasses import asdict
from importlib import resources
from pathlib import Path
from typing import Any

from cal_translator.excel.com_backend import open_excel_workbook, retry_com_call, set_com_property
from cal_translator.formats.t50_04002 import FORMAT_ID
from cal_translator.formats.t50_04002 import english_export as legacy
from cal_translator.formats.t50_04002 import english_export_v2 as v2
from cal_translator.formats.t50_04002 import english_export_v3 as v3
from cal_translator.formats.t50_04002.english_export import TranslationOperation
from cal_translator.formats.t50_04002.workbook_writer_v2 import _remove_with_retry

_NOTE_ROWS = (27, 28, 30, 31, 32, 34, 36)
_TRACE_ROWS = range(69, 72)
_PROTECTED_FORMULAS = {
    "Información!B6": "=B5",
    "Información!H37": "=I19",
    "Información!D43": "=L40+L42",
    "Información!D44": "=L41+L43",
    "Información!H43": "=L44+L46",
    "Información!H44": "=L45+L47",
}
_PROTECTED_RANGES = (
    "Información!D69:H71",
    "Resultados!C14:H20",
)


def _load_profile() -> dict[str, Any]:
    try:
        import yaml
    except ImportError as exc:
        raise RuntimeError(
            "PyYAML is required. Install the project with: python -m pip install -e ."
        ) from exc
    resource = resources.files(__package__).joinpath("temperature_american_english.yaml")
    with resource.open("r", encoding="utf-8") as handle:
        profile = yaml.safe_load(handle)
    if (
        not isinstance(profile, dict)
        or profile.get("format_id") != FORMAT_ID
        or profile.get("measurement_kind") != "temperature"
    ):
        raise RuntimeError("Invalid controlled temperature English profile.")
    return profile


def _controlled_rule(profile: dict[str, Any], name: str) -> dict[str, Any]:
    rule = (profile.get("dynamic_text") or {}).get(name) or {}
    if not isinstance(rule, dict):
        raise RuntimeError(f"Invalid controlled dynamic rule: {name}")
    return rule


def _operation(name: str, address: str, source: Any, target: Any) -> TranslationOperation:
    return TranslationOperation(
        logical_name=name,
        sheet="Información",
        address=address,
        source=str(source or ""),
        target=str(target or ""),
    )


def _translate_temperature_label(value: Any) -> str:
    source = str(value or "").strip()
    patterns = (
        (r"^Aceite$", "Oil"),
        (r"^Radiador superior\s+(\d+)$", r"Upper radiator \1"),
        (r"^Radiador inferior\s+(\d+)$", r"Lower radiator \1"),
        (r"^Temperatura ambiente\s+(\d+)$", r"Ambient temperature \1"),
    )
    for pattern, replacement in patterns:
        if re.fullmatch(pattern, source, flags=re.IGNORECASE):
            return re.sub(pattern, replacement, source, flags=re.IGNORECASE)
    raise RuntimeError(f"Uncontrolled temperature equipment label: {source!r}")


def _translate_description(value: Any) -> str:
    source = str(value or "").strip()
    match = re.fullmatch(
        r"Termocupla tipo K con registrador canal\s+(\d+)",
        source,
        flags=re.IGNORECASE,
    )
    if not match:
        raise RuntimeError(f"Uncontrolled temperature equipment description: {source!r}")
    return f"Type K thermocouple with recorder, channel {match.group(1)}"


def build_temperature_translation_plan(
    payload: dict[str, Any], profile: dict[str, Any] | None = None
) -> list[TranslationOperation]:
    profile = profile or _load_profile()
    if payload.get("format_id") != FORMAT_ID or payload.get("measurement_kind") != "temperature":
        raise RuntimeError("Temperature English export received incompatible extracted data.")

    operations = legacy._static_operations(profile)
    dynamic = profile.get("dynamic_text") or {}
    equipment = payload.get("equipment") or {}
    customer = payload.get("customer") or {}

    operations.extend(
        [
            _operation(
                "equipment.description",
                "C11",
                equipment.get("description"),
                _translate_description(equipment.get("description")),
            ),
            _operation(
                "equipment.manufacturer",
                "C13",
                equipment.get("manufacturer"),
                _controlled_rule(profile, "manufacturer").get("target"),
            ),
            _operation(
                "equipment.serial_number",
                "G13",
                equipment.get("serial_number"),
                _translate_temperature_label(equipment.get("serial_number")),
            ),
            _operation(
                "equipment.customer_identification",
                "G15",
                equipment.get("customer_identification"),
                _translate_temperature_label(equipment.get("customer_identification")),
            ),
            _operation(
                "customer.name",
                "D18",
                customer.get("name"),
                _controlled_rule(profile, "customer_name").get("target"),
            ),
            _operation(
                "customer.address",
                "D20",
                customer.get("address"),
                _controlled_rule(profile, "customer_address").get("target"),
            ),
            _operation(
                "calibration_method",
                "B26",
                payload.get("calibration_method"),
                _controlled_rule(profile, "calibration_method").get("target"),
            ),
        ]
    )

    for name, source_value in (
        ("manufacturer", equipment.get("manufacturer")),
        ("customer_name", customer.get("name")),
        ("customer_address", customer.get("address")),
        ("calibration_method", payload.get("calibration_method")),
    ):
        rule = _controlled_rule(profile, name)
        legacy._require_controlled(source_value, rule.get("source"), name)

    equipment_map = dynamic.get("traceability_equipment") or {}
    traceability = payload.get("traceability") or []
    if len(traceability) != 3:
        raise RuntimeError(f"Temperature English export requires 3 traceability rows; found {len(traceability)}.")
    for offset, entry in enumerate(traceability):
        source = str(entry.get("equipment") or "")
        if source not in equipment_map:
            raise RuntimeError(f"Uncontrolled temperature traceability equipment: {source!r}")
        operations.append(
            TranslationOperation(
                logical_name=f"traceability[{offset}].equipment",
                sheet="Información",
                address=f"B{69 + offset}",
                source=source,
                target=str(equipment_map[source]),
            )
        )

    note_map = dynamic.get("notes") or {}
    notes = payload.get("notes") or []
    if len(notes) != len(_NOTE_ROWS):
        raise RuntimeError(f"Temperature English export requires 7 notes; found {len(notes)}.")
    for row, source_value in zip(_NOTE_ROWS, notes):
        source = str(source_value)
        if source not in note_map:
            raise RuntimeError(f"Uncontrolled temperature certificate note: {source!r}")
        operations.append(
            TranslationOperation(
                logical_name=f"notes.{row}",
                sheet="Resultados",
                address=f"B{row}",
                source=source,
                target=str(note_map[source]),
            )
        )
    return operations


def _read_value(workbook: Any, location: str) -> Any:
    sheet_name, address = location.split("!", 1)
    sheet = retry_com_call(lambda: workbook.Worksheets(sheet_name))
    return retry_com_call(lambda: sheet.Range(address).Value2)


def _read_formula(workbook: Any, location: str) -> Any:
    sheet_name, address = location.split("!", 1)
    sheet = retry_com_call(lambda: workbook.Worksheets(sheet_name))
    return retry_com_call(lambda: sheet.Range(address).Formula)


def _matrix(value: Any) -> tuple[tuple[Any, ...], ...]:
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return tuple(tuple(row) for row in value)
        return (tuple(value),)
    return ((value,),)


def _protected_snapshot(workbook: Any) -> dict[str, Any]:
    ranges: dict[str, Any] = {}
    for location in _PROTECTED_RANGES:
        sheet_name, address = location.split("!", 1)
        sheet = retry_com_call(lambda sheet_name=sheet_name: workbook.Worksheets(sheet_name))
        target = retry_com_call(lambda: sheet.Range(address))
        ranges[location] = _matrix(retry_com_call(lambda: target.Value2))
    return {
        "formulas": {location: _read_formula(workbook, location) for location in _PROTECTED_FORMULAS},
        "ranges": ranges,
    }


def _apply_operations(workbook: Any, operations: list[TranslationOperation]) -> None:
    legacy._apply_operations(workbook, operations)
    info = retry_com_call(lambda: workbook.Worksheets("Información"))
    method = retry_com_call(lambda: info.Range("B26:H31"))
    set_com_property(method, "WrapText", True)
    set_com_property(method, "ShrinkToFit", False)
    set_com_property(retry_com_call(lambda: method.Font), "Size", 8.5)
    results = retry_com_call(lambda: workbook.Worksheets("Resultados"))
    set_com_property(retry_com_call(lambda: results.Range("B7:H7")), "WrapText", True)
    set_com_property(retry_com_call(lambda: results.Rows(7)), "RowHeight", 30)


def _configure_information(workbook: Any, overrides: dict[str, Any]) -> dict[str, Any]:
    info = retry_com_call(lambda: workbook.Worksheets("Información"))
    removed_shape = v2._delete_named_shape(info, "CAL_English_Laboratory_Block")
    logo_only = v3._set_logo_only_header(info)
    blocks = [
        v3._write_merged_cell_block(info, spec)
        for spec in v3.laboratory_cell_specs(overrides)
    ]
    title = v2.title_presentation(overrides)
    set_com_property(retry_com_call(lambda: info.Range("B4")), "Value2", title["certificate_number_label"])
    set_com_property(
        retry_com_call(lambda: info.Range("B6")),
        "NumberFormat",
        title["duplicate_certificate_number_format"],
    )
    trace = retry_com_call(lambda: info.Range("B69:H71"))
    set_com_property(trace, "WrapText", True)
    set_com_property(trace, "VerticalAlignment", v2._XL_CENTER)
    set_com_property(retry_com_call(lambda: info.Range("B69:C71").Font), "Size", 8.5)
    for row in _TRACE_ROWS:
        set_com_property(retry_com_call(lambda row=row: info.Rows(row)), "RowHeight", 22)
    return {
        "presentation_mode": "worksheet_cells",
        "removed_legacy_shape": removed_shape,
        "logo_only_header_configured": logo_only,
        "information_cells": blocks,
    }


def _configure_layout(workbook: Any) -> dict[str, Any]:
    info = retry_com_call(lambda: workbook.Worksheets("Información"))
    results = retry_com_call(lambda: workbook.Worksheets("Resultados"))
    set_com_property(retry_com_call(lambda: info.PageSetup), "PrintArea", "$B$1:$H$78")
    result_setup = retry_com_call(lambda: results.PageSetup)
    set_com_property(result_setup, "PrintArea", "$B$1:$H$43")
    set_com_property(result_setup, "PrintTitleRows", "")
    retry_com_call(results.ResetAllPageBreaks)
    return {"information_print_area": "B1:H78", "results_print_area": "B1:H43"}


def translate_temperature_and_export(
    source_workbook: Path,
    data_path: Path,
    output_workbook: Path,
    output_pdf: Path,
    *,
    format_id: str = FORMAT_ID,
    overwrite: bool = False,
) -> tuple[dict[str, Any], Path]:
    if format_id != FORMAT_ID:
        raise RuntimeError(f"Unsupported translation format: {format_id}. Expected {FORMAT_ID}.")
    payload = json.loads(data_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise RuntimeError("Extracted certificate JSON must contain an object.")
    if payload.get("extraction_warnings"):
        raise RuntimeError("Temperature English export requires extraction_warnings to be empty.")

    profile = _load_profile()
    operations = build_temperature_translation_plan(payload, profile)
    overrides = v2._load_overrides()
    output_workbook = output_workbook.resolve()
    output_pdf = output_pdf.resolve()
    output_workbook.parent.mkdir(parents=True, exist_ok=True)
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    for path in (output_workbook, output_pdf):
        if path.exists():
            if not overwrite:
                raise RuntimeError(f"Output already exists: {path}")
            _remove_with_retry(path)
    shutil.copy2(source_workbook.resolve(), output_workbook)

    workbook_validation: dict[str, Any]
    information_presentation: dict[str, Any] = {}
    layout: dict[str, Any] = {}
    metadata: dict[str, Any] = {}
    header_footer_changes = 0
    try:
        with open_excel_workbook(output_workbook, read_only=False) as (excel, workbook):
            protected_before = _protected_snapshot(workbook)
            _apply_operations(workbook, operations)
            header_footer_changes = v3.translate_certificate_headers_and_footers(workbook, profile)
            metadata.update(v2.configure_controlled_metadata(workbook, overrides))
            information_presentation.update(_configure_information(workbook, overrides))
            layout.update(_configure_layout(workbook))
            retry_com_call(excel.CalculateFullRebuild, attempts=20, initial_delay=0.25)
            protected_after = _protected_snapshot(workbook)
            errors: list[dict[str, Any]] = []
            if protected_after != protected_before:
                errors.append(
                    {
                        "code": "protected_data_changed",
                        "expected": protected_before,
                        "actual": protected_after,
                    }
                )
            for operation in operations:
                actual = _read_value(workbook, f"{operation.sheet}!{operation.address}")
                if legacy._normalize(actual) != legacy._normalize(operation.target):
                    errors.append(
                        {
                            "code": "translation_target_mismatch",
                            "target": f"{operation.sheet}!{operation.address}",
                            "expected": operation.target,
                            "actual": actual,
                        }
                    )
            workbook_validation = {"valid": not errors, "errors": errors}
            if not workbook_validation["valid"]:
                raise RuntimeError(
                    f"Translated temperature workbook validation failed with {len(errors)} error(s)."
                )
            retry_com_call(workbook.Save, attempts=20, initial_delay=0.25)
            legacy._export_certificate_pdf(workbook, output_pdf)
    except Exception:
        for path in (output_workbook, output_pdf):
            try:
                _remove_with_retry(path)
            except Exception:
                pass
        raise

    pdf_validation = legacy._validate_pdf(output_pdf, profile)
    if not pdf_validation["valid"]:
        raise RuntimeError(
            f"Exported temperature PDF validation failed with {len(pdf_validation['errors'])} error(s)."
        )
    report: dict[str, Any] = {
        "status": "temperature_english_export",
        "format_id": format_id,
        "measurement_kind": "temperature",
        "locale": profile.get("locale"),
        "profile_version": profile.get("profile_version"),
        "source_workbook": str(source_workbook.resolve()),
        "source_data": str(data_path.resolve()),
        "output_workbook": str(output_workbook),
        "output_pdf": str(output_pdf),
        "translation_operations": len(operations),
        "header_footer_properties_changed": header_footer_changes,
        "operations": [asdict(operation) for operation in operations],
        "source_issues": list(payload.get("source_issues") or []),
        "information_presentation": information_presentation,
        "controlled_metadata": metadata,
        "pdf_layout": layout,
        "workbook_validation": workbook_validation,
        "pdf_validation": pdf_validation,
        "valid": bool(workbook_validation["valid"] and pdf_validation["valid"]),
    }
    report_path = output_workbook.with_name(
        f"{output_workbook.stem}_translation_report.json"
    )
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    return report, report_path
