from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Iterable

from cal_translator.models import (
    CalibrationCertificate,
    CustomerData,
    EnvironmentalConditions,
    EquipmentData,
    ResultRow,
    TraceabilityEntry,
)
from cal_translator.pdf.text_extractor import PdfPageText, extract_pdf_text

_FORMAT_ID = "T50-04002"
_VALUE = r"[-+]?\d+(?:[.,]\d+)?\s*(?:m?Ω|Ω)"
_RESULT_RE = re.compile(
    rf"^(?:(?P<range>\d+(?:[.,]\d+)?\s*Ω\s*/\s*\d+(?:[.,]\d+)?\s*A)\s+)?"
    rf"(?P<specified>{_VALUE})\s+"
    rf"(?P<measured>{_VALUE})\s+"
    rf"(?P<bias>{_VALUE})\s+"
    rf"(?P<uncertainty>{_VALUE})\s+"
    r"(?P<coverage>\d+(?:[.,]\d+)?)$",
    re.IGNORECASE,
)


def _normalized_lines(text: str) -> list[str]:
    return [re.sub(r"\s+", " ", line).strip() for line in text.splitlines() if line.strip()]


def _match(text: str, pattern: str) -> str | None:
    match = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
    return match.group(1).strip() if match else None


def _value_after(lines: list[str], marker: str) -> str | None:
    try:
        index = next(i for i, line in enumerate(lines) if marker.casefold() in line.casefold())
    except StopIteration:
        return None
    for candidate in lines[index + 1 : index + 8]:
        if candidate and candidate.casefold() not in {
            "description",
            "manufacturer",
            "serial number",
            "type",
            "customer number",
            "customer",
            "street address",
            "city",
            "order number",
        }:
            return candidate
    return None


def _extract_page_one(page: PdfPageText, certificate: CalibrationCertificate) -> None:
    text = page.text
    lines = _normalized_lines(text)

    certificate.certificate_number = _match(text, r"(?:Número de certificado|Certificate number)\s*:\s*(CAL-\d+)")
    certificate.internal_code = _match(text, r"Código Interno\s*:\s*([^\r\n]+)")

    certificate.equipment = EquipmentData(
        description=_match(text, r"(HIGH CURRENT RESISTANCE METER)"),
        manufacturer=_match(text, r"\b(Tettex)\b"),
        serial_number=_match(text, r"Tettex\s+([^\s]+)"),
        model=_match(text, r"Tettex\s+[^\s]+\s+([^\s]+)"),
        customer_identification=_match(text, r"Tettex\s+[^\s]+\s+[^\s]+\s+([^\s]+)"),
    )

    certificate.customer = CustomerData(
        name=_match(text, r"(SEDT Campo de pruebas Distribución)"),
        address=_match(text, r"(km 8,5 Autopista Medellín - Costado Sur)"),
        city=_match(text, r"\n(Tenjo)\n"),
        order_number=_match(text, r"\n(No)\n"),
    )

    dates = re.findall(r"\b\d{4}-\d{2}-\d{2}\b", text)
    if len(dates) >= 3:
        certificate.reception_date = dates[0]
        certificate.calibration_date = dates[1]
        certificate.issue_date = dates[2]
    else:
        certificate.extraction_warnings.append("Expected three ISO dates on page 1.")

    environmental = re.findall(r"(\d+[.,]\d+)\s*°C\s+(\d+)\s*%", text)
    if len(environmental) >= 2:
        certificate.environmental_conditions = EnvironmentalConditions(
            maximum_temperature=f"{environmental[0][0]} °C",
            maximum_relative_humidity=f"{environmental[0][1]} %",
            minimum_temperature=f"{environmental[1][0]} °C",
            minimum_relative_humidity=f"{environmental[1][1]} %",
        )

    method_start = next((i for i, line in enumerate(lines) if line == "No"), None)
    method_end = next((i for i, line in enumerate(lines) if line.startswith("Temperatura máxima")), None)
    if method_start is not None and method_end is not None and method_start < method_end:
        certificate.calibration_method = " ".join(lines[method_start + 1 : method_end])


def _extract_traceability(page: PdfPageText) -> list[TraceabilityEntry]:
    entries: list[TraceabilityEntry] = []
    lines = _normalized_lines(page.text)
    equipment_names = [
        "Resistencia de 0,1 mΩ",
        "Resistencia de 100 mΩ",
        "Resistencia de 10 mΩ",
        "Resistencia de 1 mΩ",
        "High Power Resistance Substituter",
    ]
    row_re = re.compile(
        r"^(AEG|WLN 9|DRS-900)\s+(\d+)\s+(SET-GAT|SET-GAD)\s+([A-Z0-9-]+)\s+(\d{4}-\d{2}-\d{2})$"
    )
    matched_rows = [row_re.match(line) for line in lines]
    matched_rows = [match for match in matched_rows if match]
    for equipment, match in zip(equipment_names, matched_rows):
        entries.append(
            TraceabilityEntry(
                equipment=equipment,
                type=match.group(1),
                internal_number=match.group(2),
                calibrated_by=match.group(3),
                certificate_number=match.group(4),
                calibration_date=match.group(5),
            )
        )
    return entries


def _result_lines(pages: Iterable[PdfPageText]) -> list[str]:
    lines: list[str] = []
    for page in pages:
        if page.page_number < 3:
            continue
        lines.extend(_normalized_lines(page.text))
    return lines


def _extract_results(pages: list[PdfPageText]) -> list[ResultRow]:
    rows: list[ResultRow] = []
    channel = ""
    pending_range: str | None = None

    for line in _result_lines(pages):
        upper = line.upper()
        if upper in {"CANAL A", "CANAL B", "CANAL C"}:
            channel = upper[-1]
            continue
        match = _RESULT_RE.match(line.replace("Ω", "Ω"))
        if not match:
            continue
        if match.group("range"):
            pending_range = match.group("range")
        if not channel:
            channel = "A"
        rows.append(
            ResultRow(
                channel=channel,
                range=match.group("range") or pending_range,
                specified_value=match.group("specified"),
                average_measured_value=match.group("measured"),
                bias=match.group("bias"),
                expanded_uncertainty=match.group("uncertainty"),
                coverage_factor=match.group("coverage"),
            )
        )
    return rows


def extract_certificate(pdf_path: Path) -> CalibrationCertificate:
    pages = extract_pdf_text(pdf_path)
    certificate = CalibrationCertificate(format_id=_FORMAT_ID, source_pages=len(pages))
    if not pages:
        certificate.extraction_warnings.append("PDF contains no pages.")
        return certificate

    _extract_page_one(pages[0], certificate)
    if len(pages) >= 2:
        certificate.traceability = _extract_traceability(pages[1])
    certificate.results = _extract_results(pages)

    if certificate.certificate_number is None:
        certificate.extraction_warnings.append("Certificate number was not extracted.")
    if len(certificate.traceability) != 5:
        certificate.extraction_warnings.append(
            f"Expected 5 traceability entries, extracted {len(certificate.traceability)}."
        )
    if len(certificate.results) != 72:
        certificate.extraction_warnings.append(
            f"Expected 72 result rows for CAL-13078, extracted {len(certificate.results)}."
        )
    return certificate


def write_extraction(certificate: CalibrationCertificate, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    certificate_number = certificate.certificate_number or "certificate"
    output_path = output_dir / f"{certificate_number}_extracted.json"
    output_path.write_text(
        json.dumps(certificate.to_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return output_path
