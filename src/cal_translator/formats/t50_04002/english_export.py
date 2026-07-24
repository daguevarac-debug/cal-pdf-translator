from __future__ import annotations

import json
import re
import shutil
from dataclasses import asdict, dataclass
from importlib import resources
from pathlib import Path
from typing import Any

from cal_translator.excel.com_backend import (
    open_excel_workbook,
    retry_com_call,
    set_com_property,
)
from cal_translator.formats.t50_04002 import FORMAT_ID
from cal_translator.formats.t50_04002.workbook_writer_v2 import _remove_with_retry

_XL_TYPE_PDF = 0
_XL_QUALITY_STANDARD = 0
_XL_SHEET_HIDDEN = 0

_FORMULAS = {
    "Información!B6": "=B5",
    "Información!H37": "=I19",
    "Información!D43": "=L40+L42",
    "Información!D44": "=L41+L43",
    "Información!H43": "=L44+L46",
    "Información!H44": "=L45+L47",
}
_PROTECTED_SCALARS = (
    "Información!B5",
    "Información!B7",
    "Información!C11",
    "Información!C13",
    "Información!G13",
    "Información!C15",
    "Información!G15",
    "Información!D21",
    "Información!D22",
    "Información!D46",
    "Información!H46",
    "Información!D48",
    "Información!L40",
    "Información!L41",
    "Información!L42",
    "Información!L43",
    "Información!L44",
    "Información!L45",
    "Información!L46",
    "Información!L47",
)
_PROTECTED_RANGES = (
    "Información!D69:H73",
    "Resultados!B10:L33",
    "Resultados!B35:L58",
    "Resultados!B60:L83",
)
_NOTE_ROWS = (86, 87, 89, 90, 91, 93, 95)


@dataclass(slots=True)
class TranslationOperation:
    logical_name: str
    sheet: str
    address: str
    source: str
    target: str


def _load_profile() -> dict[str, Any]:
    try:
        import yaml
    except ImportError as exc:
        raise RuntimeError(
            "PyYAML is required. Install the project with: python -m pip install -e ."
        ) from exc

    resource = resources.files(__package__).joinpath("american_english.yaml")
    with resource.open("r", encoding="utf-8") as handle:
        profile = yaml.safe_load(handle)
    if not isinstance(profile, dict) or profile.get("format_id") != FORMAT_ID:
        raise RuntimeError("Invalid controlled American English profile for T50-04002.")
    return profile


def _normalize(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def _require_controlled(source: Any, expected: Any, logical_name: str) -> None:
    if _normalize(source) != _normalize(expected):
        raise RuntimeError(
            f"Uncontrolled source text for {logical_name}. "
            f"Expected {expected!r}, found {source!r}."
        )


def _static_operations(profile: dict[str, Any]) -> list[TranslationOperation]:
    operations: list[TranslationOperation] = []
    seen: set[tuple[str, str]] = set()
    for index, item in enumerate(profile.get("static_cells") or []):
        sheet = str(item["sheet"])
        address = str(item["address"])
        key = (sheet, address)
        if key in seen:
            raise RuntimeError(f"Duplicate controlled translation target: {sheet}!{address}")
        seen.add(key)
        operations.append(
            TranslationOperation(
                logical_name=f"static_cells[{index}]",
                sheet=sheet,
                address=address,
                source=str(item.get("source") or ""),
                target=str(item.get("target") or ""),
            )
        )

    for name, item in (profile.get("static_paragraphs") or {}).items():
        sheet = str(item["sheet"])
        address = str(item["address"])
        key = (sheet, address)
        if key in seen:
            raise RuntimeError(f"Duplicate controlled translation target: {sheet}!{address}")
        seen.add(key)
        operations.append(
            TranslationOperation(
                logical_name=f"static_paragraphs.{name}",
                sheet=sheet,
                address=address,
                source=str(item.get("source") or ""),
                target=str(item.get("target") or ""),
            )
        )
    return operations


def _dynamic_operations(payload: dict[str, Any], profile: dict[str, Any]) -> list[TranslationOperation]:
    dynamic = profile.get("dynamic_text") or {}
    customer = payload.get("customer") or {}
    operations: list[TranslationOperation] = []

    for logical_name, address, source_value, rule_name in (
        ("customer.name", "D18", customer.get("name"), "customer_name"),
        ("customer.address", "D20", customer.get("address"), "customer_address"),
        ("calibration_method", "B26", payload.get("calibration_method"), "calibration_method"),
    ):
        rule = dynamic.get(rule_name) or {}
        _require_controlled(source_value, rule.get("source"), logical_name)
        operations.append(
            TranslationOperation(
                logical_name=logical_name,
                sheet="Información",
                address=address,
                source=str(source_value or ""),
                target=str(rule.get("target") or ""),
            )
        )

    equipment_map = dynamic.get("traceability_equipment") or {}
    traceability = payload.get("traceability") or []
    if len(traceability) != 5:
        raise RuntimeError(f"Controlled English export requires 5 traceability rows; found {len(traceability)}.")
    for offset, entry in enumerate(traceability):
        source = str(entry.get("equipment") or "")
        if source not in equipment_map:
            raise RuntimeError(f"Uncontrolled traceability equipment text: {source!r}")
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
        raise RuntimeError(
            f"Controlled English export requires {len(_NOTE_ROWS)} notes; found {len(notes)}."
        )
    for row, source in zip(_NOTE_ROWS, notes):
        source = str(source)
        if source not in note_map:
            raise RuntimeError(f"Uncontrolled certificate note: {source!r}")
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


def build_translation_plan(payload: dict[str, Any]) -> list[TranslationOperation]:
    if payload.get("format_id") != FORMAT_ID:
        raise RuntimeError(
            f"Unsupported extracted-data format: {payload.get('format_id')}. Expected {FORMAT_ID}."
        )
    profile = _load_profile()
    return [*_static_operations(profile), *_dynamic_operations(payload, profile)]


def _get_sheet(workbook: Any, name: str) -> Any:
    return retry_com_call(lambda: workbook.Worksheets(name))


def _get_range(workbook: Any, location: str) -> Any:
    sheet_name, address = location.split("!", 1)
    sheet = _get_sheet(workbook, sheet_name)
    return retry_com_call(lambda: sheet.Range(address))


def _read_value(workbook: Any, location: str) -> Any:
    cell = _get_range(workbook, location)
    return retry_com_call(lambda: cell.Value2)


def _read_formula(workbook: Any, location: str) -> Any:
    cell = _get_range(workbook, location)
    return retry_com_call(lambda: cell.Formula)


def _matrix(value: Any) -> tuple[tuple[Any, ...], ...]:
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return tuple(tuple(row) for row in value)
        return (tuple(value),)
    return ((value,),)


def _protected_snapshot(workbook: Any) -> dict[str, Any]:
    snapshot: dict[str, Any] = {
        "formulas": {location: _read_formula(workbook, location) for location in _FORMULAS},
        "scalars": {location: _read_value(workbook, location) for location in _PROTECTED_SCALARS},
        "ranges": {},
    }
    for location in _PROTECTED_RANGES:
        target = _get_range(workbook, location)
        snapshot["ranges"][location] = _matrix(retry_com_call(lambda target=target: target.Value2))
    return snapshot


def _apply_operations(workbook: Any, operations: list[TranslationOperation]) -> None:
    for operation in operations:
        sheet = _get_sheet(workbook, operation.sheet)
        cell = retry_com_call(lambda sheet=sheet, operation=operation: sheet.Range(operation.address))
        actual = retry_com_call(lambda cell=cell: cell.Value2)
        _require_controlled(actual, operation.source, operation.logical_name)
        retry_com_call(lambda cell=cell, operation=operation: setattr(cell, "Value2", operation.target))

    info = _get_sheet(workbook, "Información")
    method = retry_com_call(lambda: info.Range("B26:H31"))
    set_com_property(method, "WrapText", True)
    set_com_property(method, "ShrinkToFit", False)
    font = retry_com_call(lambda: method.Font)
    set_com_property(font, "Size", 8.5)

    result_headers = retry_com_call(lambda: _get_sheet(workbook, "Resultados").Range("B7:L7"))
    set_com_property(result_headers, "WrapText", True)
    set_com_property(retry_com_call(lambda: _get_sheet(workbook, "Resultados").Rows(7)), "RowHeight", 30)


def _translate_headers_and_footers(workbook: Any, profile: dict[str, Any]) -> int:
    replacements = [
        (str(item["source"]), str(item["target"]))
        for item in profile.get("header_footer_replacements") or []
    ]
    changed = 0
    worksheet_count = int(retry_com_call(lambda: workbook.Worksheets.Count))
    properties = (
        "LeftHeader",
        "CenterHeader",
        "RightHeader",
        "LeftFooter",
        "CenterFooter",
        "RightFooter",
    )
    for index in range(1, worksheet_count + 1):
        sheet = retry_com_call(lambda index=index: workbook.Worksheets(index))
        page_setup = retry_com_call(lambda sheet=sheet: sheet.PageSetup)
        for name in properties:
            current = str(retry_com_call(lambda page_setup=page_setup, name=name: getattr(page_setup, name)) or "")
            updated = current
            for source, target in replacements:
                updated = updated.replace(source, target)
            if updated != current:
                set_com_property(page_setup, name, updated)
                changed += 1
    return changed


def _validate_translated_workbook(
    workbook: Any,
    operations: list[TranslationOperation],
    protected_before: dict[str, Any],
) -> dict[str, Any]:
    errors: list[dict[str, Any]] = []
    for operation in operations:
        actual = _read_value(workbook, f"{operation.sheet}!{operation.address}")
        if _normalize(actual) != _normalize(operation.target):
            errors.append(
                {
                    "code": "translation_target_mismatch",
                    "target": f"{operation.sheet}!{operation.address}",
                    "expected": operation.target,
                    "actual": actual,
                }
            )

    protected_after = _protected_snapshot(workbook)
    if protected_after != protected_before:
        errors.append(
            {
                "code": "protected_data_changed",
                "target": "formulas, scalar certificate values, traceability metadata or result data",
                "expected": protected_before,
                "actual": protected_after,
            }
        )
    return {"valid": not errors, "errors": errors}


def validate_pdf_text(text: str, page_count: int, profile: dict[str, Any] | None = None) -> dict[str, Any]:
    profile = profile or _load_profile()
    rules = profile.get("pdf_validation") or {}
    normalized = re.sub(r"\s+", " ", text).casefold()
    missing = [
        phrase
        for phrase in rules.get("required_phrases") or []
        if _normalize(phrase) not in normalized
    ]
    forbidden = [
        phrase
        for phrase in rules.get("forbidden_phrases") or []
        if _normalize(phrase) in normalized
    ]
    expected_pages = int(rules.get("expected_pages") or 0)
    errors: list[dict[str, Any]] = []
    if page_count != expected_pages:
        errors.append(
            {
                "code": "pdf_page_count",
                "expected": expected_pages,
                "actual": page_count,
            }
        )
    if missing:
        errors.append({"code": "missing_english_phrases", "phrases": missing})
    if forbidden:
        errors.append({"code": "remaining_spanish_phrases", "phrases": forbidden})
    return {
        "valid": not errors,
        "page_count": page_count,
        "missing_required_phrases": missing,
        "remaining_forbidden_phrases": forbidden,
        "errors": errors,
    }


def _validate_pdf(pdf_path: Path, profile: dict[str, Any]) -> dict[str, Any]:
    try:
        import fitz
    except ImportError as exc:
        raise RuntimeError("PyMuPDF is required to validate the exported PDF.") from exc

    if not pdf_path.exists() or pdf_path.stat().st_size == 0:
        raise RuntimeError(f"Excel did not create a valid PDF file: {pdf_path}")
    with fitz.open(pdf_path) as document:
        text = "\n".join(page.get_text("text") for page in document)
        return validate_pdf_text(text, len(document), profile)


def _export_certificate_pdf(workbook: Any, pdf_path: Path) -> None:
    hidden_states: dict[str, Any] = {}
    try:
        for name in ("Administración", "Control de cambios"):
            sheet = _get_sheet(workbook, name)
            hidden_states[name] = retry_com_call(lambda sheet=sheet: sheet.Visible)
            set_com_property(sheet, "Visible", _XL_SHEET_HIDDEN)
        retry_com_call(
            lambda: workbook.ExportAsFixedFormat(
                Type=_XL_TYPE_PDF,
                Filename=str(pdf_path),
                Quality=_XL_QUALITY_STANDARD,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            ),
            attempts=20,
            initial_delay=0.25,
        )
    finally:
        for name, state in hidden_states.items():
            try:
                set_com_property(_get_sheet(workbook, name), "Visible", state)
            except Exception:  # noqa: BLE001
                pass


def translate_and_export(
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
    if not source_workbook.exists():
        raise FileNotFoundError(f"Expanded workbook not found: {source_workbook}")
    if not data_path.exists():
        raise FileNotFoundError(f"Extracted certificate JSON not found: {data_path}")

    payload = json.loads(data_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise RuntimeError("Extracted certificate JSON must contain an object.")
    if payload.get("extraction_warnings"):
        raise RuntimeError("English export requires extraction_warnings to be empty.")

    output_workbook = output_workbook.resolve()
    output_pdf = output_pdf.resolve()
    if output_workbook.suffix.lower() != ".xlsx":
        raise RuntimeError("English workbook output must use the .xlsx extension.")
    if output_pdf.suffix.lower() != ".pdf":
        raise RuntimeError("English PDF output must use the .pdf extension.")
    if source_workbook.resolve() == output_workbook:
        raise RuntimeError("The English output workbook must not overwrite the expanded source workbook.")

    output_workbook.parent.mkdir(parents=True, exist_ok=True)
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    for path in (output_workbook, output_pdf):
        if path.exists():
            if not overwrite:
                raise RuntimeError(f"Output already exists: {path}")
            _remove_with_retry(path)

    profile = _load_profile()
    operations = build_translation_plan(payload)
    shutil.copy2(source_workbook.resolve(), output_workbook)

    workbook_validation: dict[str, Any]
    header_footer_changes = 0
    try:
        with open_excel_workbook(output_workbook, read_only=False) as (excel, workbook):
            if bool(retry_com_call(lambda: workbook.ReadOnly)):
                raise RuntimeError("Copied English workbook unexpectedly opened as read-only.")
            protected_before = _protected_snapshot(workbook)
            _apply_operations(workbook, operations)
            header_footer_changes = _translate_headers_and_footers(workbook, profile)
            retry_com_call(excel.CalculateFullRebuild, attempts=20, initial_delay=0.25)
            workbook_validation = _validate_translated_workbook(
                workbook,
                operations,
                protected_before,
            )
            if not workbook_validation["valid"]:
                raise RuntimeError(
                    f"Translated workbook validation failed with {len(workbook_validation['errors'])} error(s)."
                )
            retry_com_call(workbook.Save, attempts=20, initial_delay=0.25)
            _export_certificate_pdf(workbook, output_pdf)
    except Exception:
        for path in (output_workbook, output_pdf):
            try:
                _remove_with_retry(path)
            except Exception:  # noqa: BLE001
                pass
        raise

    pdf_validation = _validate_pdf(output_pdf, profile)
    if not pdf_validation["valid"]:
        raise RuntimeError(
            f"Exported PDF validation failed with {len(pdf_validation['errors'])} error(s)."
        )

    report: dict[str, Any] = {
        "status": "english_export_prototype",
        "format_id": format_id,
        "locale": profile.get("locale"),
        "profile_version": profile.get("profile_version"),
        "source_workbook": str(source_workbook.resolve()),
        "source_data": str(data_path.resolve()),
        "output_workbook": str(output_workbook),
        "output_pdf": str(output_pdf),
        "translation_operations": len(operations),
        "header_footer_properties_changed": header_footer_changes,
        "operations": [asdict(operation) for operation in operations],
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
