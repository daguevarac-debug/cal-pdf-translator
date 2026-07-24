from __future__ import annotations

import re
from pathlib import Path

from cal_translator.formats.t50_04002 import temperature_extractor as legacy
from cal_translator.models import CalibrationCertificate, EnvironmentalConditions
from cal_translator.pdf.text_extractor import PdfPageText

_ENVIRONMENT_WARNING = "Expected two environmental condition rows on page 1."
_ORIGINAL_EXTRACT_PAGE_ONE = legacy._extract_page_one
_ORIGINAL_EXTRACT_CERTIFICATE = legacy.extract_temperature_certificate


def extract_labeled_environmental_conditions(text: str) -> EnvironmentalConditions | None:
    """Extract environmental values when labels and values occupy separate PDF lines."""
    patterns = {
        "maximum_temperature": r"Temperatura\s+m[aá]xima\s*\r?\n\s*(\d+[.,]\d+)\s*°C",
        "minimum_temperature": r"Temperatura\s+m[ií]nima\s*\r?\n\s*(\d+[.,]\d+)\s*°C",
        "maximum_relative_humidity": r"Humedad\s+Relativa\s+m[aá]xima\s*\r?\n\s*(\d+[.,]?\d*)\s*%",
        "minimum_relative_humidity": r"Humedad\s+Relativa\s+m[ií]nima\s*\r?\n\s*(\d+[.,]?\d*)\s*%",
    }
    values: dict[str, str] = {}
    for field, pattern in patterns.items():
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if not match:
            return None
        unit = "°C" if "temperature" in field else "%"
        values[field] = f"{match.group(1)} {unit}"
    return EnvironmentalConditions(**values)


def _extract_page_one(page: PdfPageText, certificate: CalibrationCertificate) -> None:
    _ORIGINAL_EXTRACT_PAGE_ONE(page, certificate)
    environmental = extract_labeled_environmental_conditions(page.text)
    if environmental is None:
        return
    certificate.environmental_conditions = environmental
    certificate.extraction_warnings = [
        warning
        for warning in certificate.extraction_warnings
        if warning != _ENVIRONMENT_WARNING
    ]


def extract_temperature_certificate(pdf_path: Path) -> CalibrationCertificate:
    """Run the temperature extractor with split-row environmental parsing enabled."""
    original_page_one = legacy._extract_page_one
    legacy._extract_page_one = _extract_page_one
    try:
        return _ORIGINAL_EXTRACT_CERTIFICATE(pdf_path)
    finally:
        legacy._extract_page_one = original_page_one


def install_patch() -> None:
    """Expose the corrected extractor through the legacy module used by the batch CLI."""
    legacy.extract_temperature_certificate = extract_temperature_certificate


is_temperature_certificate = legacy.is_temperature_certificate
