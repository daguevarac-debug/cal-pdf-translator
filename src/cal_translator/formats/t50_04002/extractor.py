from __future__ import annotations

import json
import re
from decimal import Decimal, InvalidOperation
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
_EXPECTED_RESULT_ROWS = 72
_ROWS_PER_CHANNEL = 24
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


def _decimal(value: str) -> Decimal:
    return Decimal(value.replace(",", "."))


def _extract_page_one(page: PdfPageText, certificate: CalibrationCertificate) -> None:
    text = page.text
    lines = _normalized_lines(text)

    certificate.certificate_number = _match(
        text,
        r"(?:Número de certificado|Certificate number)\s*:\s*(CAL-\d+)",
    )
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
        # In this controlled layout, PyMuPDF emits calibration and issue dates
        # before the reception date because of their drawing positions.
        certificate.calibration_date = dates[0]
        certificate.issue_date = dates[1]
        certificate.reception_date = dates[2]
    else:
        certificate.extraction_warnings.append("Expected three ISO dates on page 1.")

    environmental = re.findall(r"(\d+[.,]\d+)\s*°C\s+(\d+[.,]?\d*)\s*%", text)
    if len(environmental) >= 2:
        try:
            temperatures = sorted(environmental, key=lambda item: _decimal(item[0]))
            humidities = sorted(environmental, key=lambda item: _decimal(item[1]))
            certificate.environmental_conditions = EnvironmentalConditions(
                maximum_temperature=f"{temperatures[-1][0]} °C",
                minimum_temperature=f"{temperatures[0][0]} °C",
                maximum_relative_humidity=f"{humidities[-1][1]} %",
                minimum_relative_humidity=f"{humidities[0][1]} %",
            )
        except InvalidOperation:
            certificate.extraction_warnings.append("Environmental values could not be ordered numerically.")
    else:
        certificate.extraction_warnings.append("Expected two environmental condition rows on page 1.")

    method_start = next((i for i, line in enumerate(lines) if line == "No"), None)
    method_end = next((i for i, line in enumerate(lines) if line.startswith("Temperatura máxima")), None)
    if method_start is not None and method_end is not None and method_start < method_end:
        certificate.calibration_method = " ".join(lines[method_start + 1 : method_end])
    else:
        certificate.extraction_warnings.append("Calibration method block was not delimited.")


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
    matched_rows = [match for line in lines if (match := row_re.match(line))]
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
        if page.page_number >= 3:
            lines.extend(_normalized_lines(page.text))
    return lines


def _channel_for_index(index: int) -> str:
    channel_index = index // _ROWS_PER_CHANNEL
    return ("A", "B", "C")[channel_index] if channel_index < 3 else "UNKNOWN"


def _extract_results(pages: list[PdfPageText]) -> list[ResultRow]:
    parsed: list[dict[str, str | None]] = []
    pending_range: str | None = None

    for line in _result_lines(pages):
        match = _RESULT_RE.match(line.replace("Ω", "Ω"))
        if not match:
            continue
        if match.group("range"):
            pending_range = match.group("range")
        parsed.append(
            {
                "range": match.group("range") or pending_range,
                "specified": match.group("specified"),
                "measured": match.group("measured"),
                "bias": match.group("bias"),
                "uncertainty": match.group("uncertainty"),
                "coverage": match.group("coverage"),
            }
        )

    rows: list[ResultRow] = []
    for index, values in enumerate(parsed):
        rows.append(
            ResultRow(
                channel=_channel_for_index(index),
                range=values["range"],
                specified_value=str(values["specified"]),
                average_measured_value=str(values["measured"]),
                bias=str(values["bias"]),
                expanded_uncertainty=str(values["uncertainty"]),
                coverage_factor=str(values["coverage"]),
            )
        )
    return rows


def _extract_notes(page: PdfPageText) -> list[str]:
    lines = _normalized_lines(page.text)
    notes: dict[int, list[str]] = {}
    current_number: int | None = None

    for line in lines:
        match = re.match(r"^(\d+)\.\s*(.*)$", line)
        if match:
            current_number = int(match.group(1))
            notes[current_number] = [match.group(2)]
            continue
        if current_number is not None:
            if line.startswith("** Fin del certificado") or line.startswith("Unrestricted Formato Base"):
                current_number = None
            else:
                notes[current_number].append(line)

    return [f"{number}. {' '.join(notes[number]).strip()}" for number in sorted(notes)]


def extract_certificate(pdf_path: Path) -> CalibrationCertificate:
    pages = extract_pdf_text(pdf_path)
    certificate = CalibrationCertificate(format_id=_FORMAT_ID, source_pages=len(pages))
    if not pages:
        certificate.extraction_warnings.append("PDF contains no pages.")
        return certificate

    source_format = _match("\n".join(page.text for page in pages), r"Código Documento\s*:\s*(T50-\d+)")
    if source_format != _FORMAT_ID:
        certificate.extraction_warnings.append(
            f"Expected source format {_FORMAT_ID}, found {source_format or 'unknown'}."
        )

    _extract_page_one(pages[0], certificate)
    if len(pages) >= 2:
        certificate.traceability = _extract_traceability(pages[1])
    certificate.results = _extract_results(pages)
    certificate.notes = _extract_notes(pages[-1])

    if certificate.certificate_number is None:
        certificate.extraction_warnings.append("Certificate number was not extracted.")
    if len(certificate.traceability) != 5:
        certificate.extraction_warnings.append(
            f"Expected 5 traceability entries, extracted {len(certificate.traceability)}."
        )
    if len(certificate.results) != _EXPECTED_RESULT_ROWS:
        certificate.extraction_warnings.append(
            f"Expected {_EXPECTED_RESULT_ROWS} result rows for CAL-13078, extracted {len(certificate.results)}."
        )

    channel_counts = {
        channel: sum(row.channel == channel for row in certificate.results)
        for channel in ("A", "B", "C")
    }
    for channel, count in channel_counts.items():
        if count != _ROWS_PER_CHANNEL:
            certificate.extraction_warnings.append(
                f"Expected {_ROWS_PER_CHANNEL} rows for channel {channel}, extracted {count}."
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
