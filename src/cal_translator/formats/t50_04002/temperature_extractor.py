from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

from cal_translator.formats.t50_04002.extractor import (
    _extract_notes,
    _match,
    _normalize_line,
)
from cal_translator.models import (
    CalibrationCertificate,
    CustomerData,
    EnvironmentalConditions,
    EquipmentData,
    ResultRow,
    TraceabilityEntry,
)
from cal_translator.pdf.text_extractor import PdfPageText, PdfWord, extract_pdf_text

FORMAT_ID = "T50-04002"
MEASUREMENT_KIND = "temperature"
_EXPECTED_PAGES = 3
_EXPECTED_TRACEABILITY = 3
_EXPECTED_RESULTS = 7
_TEMP_VALUE = r"[-+]?\d+(?:[.,]\d+)?\s*°C"
_RESULT_RE = re.compile(
    rf"(?P<specified>{_TEMP_VALUE})\s+"
    rf"(?P<measured>{_TEMP_VALUE})\s+"
    rf"(?P<bias>{_TEMP_VALUE})\s+"
    rf"(?P<maximum>{_TEMP_VALUE})\s+"
    rf"(?P<uncertainty>{_TEMP_VALUE})\s+"
    r"(?P<coverage>\d+(?:[.,]\d+)?)\s*$",
    re.IGNORECASE,
)


def is_temperature_certificate(pages: Iterable[PdfPageText]) -> bool:
    text = "\n".join(page.text for page in pages)
    return "medición de temperatura" in text.casefold() or "termocupla tipo k" in text.casefold()


def _cluster_rows(words: list[PdfWord], *, y_min: float, y_max: float, tolerance: float = 3.2) -> list[list[PdfWord]]:
    selected = [word for word in words if y_min <= word.vertical_center <= y_max]
    selected.sort(key=lambda word: (word.vertical_center, word.x0))
    rows: list[list[PdfWord]] = []
    centers: list[float] = []
    for word in selected:
        if not rows or abs(word.vertical_center - centers[-1]) > tolerance:
            rows.append([word])
            centers.append(word.vertical_center)
            continue
        rows[-1].append(word)
        centers[-1] = sum(item.vertical_center for item in rows[-1]) / len(rows[-1])
    return [sorted(row, key=lambda word: word.x0) for row in rows]


def _join(words: Iterable[PdfWord]) -> str:
    return _normalize_line(" ".join(word.text for word in sorted(words, key=lambda word: word.x0)))


def _box_text(page: PdfPageText, *, x_min: float, x_max: float, y_min: float, y_max: float) -> str | None:
    selected = [
        word
        for word in page.words
        if x_min <= (word.x0 + word.x1) / 2.0 <= x_max
        and y_min <= word.vertical_center <= y_max
    ]
    if not selected:
        return None
    rows = _cluster_rows(selected, y_min=y_min, y_max=y_max)
    value = " ".join(_join(row) for row in rows if row).strip()
    return value or None


def _decimal(value: str) -> Decimal:
    return Decimal(value.replace(",", "."))


def _extract_method(text: str) -> str | None:
    match = re.search(
        r"(La calibración de termómetros.*?THERMOMETERS,\s*CONTACT,\s*DIRECT\s*READING:\s*CALIBRATION)",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    return _normalize_line(match.group(1)) if match else None


def _extract_page_one(page: PdfPageText, certificate: CalibrationCertificate) -> None:
    text = page.text
    certificate.certificate_number = _match(
        text,
        r"(?:Número de certificado|Certificate number)\s*:\s*(CAL-\d+)",
    )
    certificate.internal_code = _match(text, r"Código Interno\s*:\s*([^\r\n]+)")
    certificate.equipment = EquipmentData(
        description=_box_text(page, x_min=125, x_max=380, y_min=248, y_max=274),
        manufacturer=_box_text(page, x_min=125, x_max=250, y_min=276, y_max=300),
        serial_number=_box_text(page, x_min=395, x_max=535, y_min=276, y_max=300),
        model=_box_text(page, x_min=125, x_max=250, y_min=305, y_max=330),
        customer_identification=_box_text(page, x_min=395, x_max=535, y_min=305, y_max=330),
    )
    certificate.customer = CustomerData(
        name=_box_text(page, x_min=185, x_max=410, y_min=350, y_max=380),
        address=_box_text(page, x_min=185, x_max=430, y_min=380, y_max=405),
        city=_box_text(page, x_min=185, x_max=300, y_min=402, y_max=421),
        order_number=_box_text(page, x_min=185, x_max=300, y_min=421, y_max=443),
    )

    dates = re.findall(r"\b\d{4}-\d{2}-\d{2}\b", text)
    if len(dates) >= 3:
        certificate.reception_date = dates[0]
        certificate.calibration_date = dates[1]
        certificate.issue_date = dates[2]
    else:
        certificate.extraction_warnings.append("Expected reception, calibration and issue dates on page 1.")

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

    certificate.calibration_method = _extract_method(text)
    if not certificate.calibration_method:
        certificate.extraction_warnings.append("Temperature calibration method block was not extracted.")


def _column_text(row: list[PdfWord], x_min: float, x_max: float) -> str:
    return _join(
        word
        for word in row
        if x_min <= (word.x0 + word.x1) / 2.0 < x_max
    )


def _extract_traceability(page: PdfPageText, certificate: CalibrationCertificate) -> list[TraceabilityEntry]:
    entries: list[TraceabilityEntry] = []
    for row in _cluster_rows(page.words, y_min=418, y_max=458, tolerance=4.0):
        internal_number = _column_text(row, 260, 320)
        if not re.fullmatch(r"\d{8}", internal_number):
            continue
        entry = TraceabilityEntry(
            equipment=_column_text(row, 55, 200),
            type=_column_text(row, 200, 260),
            internal_number=internal_number,
            calibrated_by=_column_text(row, 320, 405),
            certificate_number=_column_text(row, 405, 475),
            calibration_date=_column_text(row, 475, 540),
        )
        if "REF" in entry.calibration_date.upper():
            certificate.source_issues.append(
                f"Source traceability date contains an Excel reference error for {entry.equipment}: {entry.calibration_date}."
            )
        entries.append(entry)
    return entries


def _extract_result_metadata(page: PdfPageText) -> dict[str, str]:
    return {
        "measurement_title": "Medición de Temperatura",
        "accuracy": _box_text(page, x_min=140, x_max=205, y_min=268, y_max=280) or "",
        "from": _box_text(page, x_min=140, x_max=200, y_min=281, y_max=296) or "",
        "to": _box_text(page, x_min=255, x_max=325, y_min=281, y_max=296) or "",
    }


def _extract_results(page: PdfPageText) -> list[ResultRow]:
    results: list[ResultRow] = []
    for row in _cluster_rows(page.words, y_min=294, y_max=386, tolerance=3.2):
        line = _join(row)
        match = _RESULT_RE.search(line)
        if not match:
            continue
        results.append(
            ResultRow(
                channel="TEMPERATURE",
                range=None,
                specified_value=_normalize_line(match.group("specified")),
                average_measured_value=_normalize_line(match.group("measured")),
                bias=_normalize_line(match.group("bias")),
                maximum_permissible_error=_normalize_line(match.group("maximum")),
                expanded_uncertainty=_normalize_line(match.group("uncertainty")),
                coverage_factor=match.group("coverage"),
            )
        )
    return results


def extract_temperature_certificate(pdf_path: Path) -> CalibrationCertificate:
    pages = extract_pdf_text(pdf_path)
    certificate = CalibrationCertificate(
        format_id=FORMAT_ID,
        measurement_kind=MEASUREMENT_KIND,
        source_pages=len(pages),
    )
    if not pages:
        certificate.extraction_warnings.append("PDF contains no pages.")
        return certificate

    source_format = _match("\n".join(page.text for page in pages), r"Código Documento\s*:\s*(T50-\d+)")
    if source_format != FORMAT_ID:
        certificate.extraction_warnings.append(
            f"Expected source format {FORMAT_ID}, found {source_format or 'unknown'}."
        )
    if len(pages) != _EXPECTED_PAGES:
        certificate.extraction_warnings.append(
            f"Expected {_EXPECTED_PAGES} pages for a temperature certificate, found {len(pages)}."
        )

    _extract_page_one(pages[0], certificate)
    if len(pages) >= 2:
        certificate.traceability = _extract_traceability(pages[1], certificate)
    if len(pages) >= 3:
        certificate.result_metadata = _extract_result_metadata(pages[2])
        certificate.results = _extract_results(pages[2])
        certificate.notes = _extract_notes(pages[2])

    if not certificate.certificate_number:
        certificate.extraction_warnings.append("Certificate number was not extracted.")
    if len(certificate.traceability) != _EXPECTED_TRACEABILITY:
        certificate.extraction_warnings.append(
            f"Expected {_EXPECTED_TRACEABILITY} traceability entries, extracted {len(certificate.traceability)}."
        )
    if len(certificate.results) != _EXPECTED_RESULTS:
        certificate.extraction_warnings.append(
            f"Expected {_EXPECTED_RESULTS} temperature result rows, extracted {len(certificate.results)}."
        )
    if len(certificate.notes) != 7:
        certificate.extraction_warnings.append(
            f"Expected 7 certificate notes, extracted {len(certificate.notes)}."
        )
    required_metadata = ("accuracy", "from", "to")
    missing_metadata = [key for key in required_metadata if not certificate.result_metadata.get(key)]
    if missing_metadata:
        certificate.extraction_warnings.append(
            "Missing temperature result metadata: " + ", ".join(missing_metadata) + "."
        )
    return certificate
