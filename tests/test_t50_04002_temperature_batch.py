from cal_translator.formats.t50_04002.temperature_english_export import (
    _load_profile,
    _translate_description,
    _translate_temperature_label,
)
from cal_translator.formats.t50_04002.temperature_extractor import (
    is_temperature_certificate,
)
from cal_translator.formats.t50_04002.temperature_workbook_writer import (
    temperature_result_matrix,
)
from cal_translator.formats.t50_04002.temperature_workbook_writer_v2 import (
    _percent_equivalent,
)
from cal_translator.pdf.text_extractor import PdfPageText


def test_temperature_certificate_is_detected_from_measurement_title() -> None:
    pages = [
        PdfPageText(
            page_number=1,
            text="Código Documento: T50-04002\nMedición de Temperatura",
        )
    ]
    assert is_temperature_certificate(pages) is True


def test_temperature_equipment_text_is_translated_by_controlled_patterns() -> None:
    assert (
        _translate_description("Termocupla tipo K con registrador canal 2")
        == "Type K thermocouple with recorder, channel 2"
    )
    assert _translate_temperature_label("Aceite") == "Oil"
    assert _translate_temperature_label("Radiador superior 2") == "Upper radiator 2"
    assert _translate_temperature_label("Radiador inferior 1") == "Lower radiator 1"
    assert _translate_temperature_label("Temperatura Ambiente 4") == "Ambient temperature 4"


def test_temperature_result_matrix_preserves_source_values() -> None:
    payload = {
        "results": [
            {
                "specified_value": "70,045 °C",
                "average_measured_value": "49,59 °C",
                "bias": "-20,5 °C",
                "maximum_permissible_error": "0,700 °C",
                "expanded_uncertainty": "90,5 °C",
                "coverage_factor": "4,5",
            }
        ]
    }
    assert temperature_result_matrix(payload) == (
        ("70,045 °C", "49,59 °C", "-20,5 °C", "0,700 °C", "90,5 °C", "4,5"),
    )


def test_temperature_profile_requires_three_pdf_pages() -> None:
    profile = _load_profile()
    assert profile["measurement_kind"] == "temperature"
    assert profile["pdf_validation"]["expected_pages"] == 3


def test_percentage_text_matches_excel_internal_fraction() -> None:
    assert _percent_equivalent("1%", 0.01) is True
    assert _percent_equivalent("0,5%", 0.005) is True
    assert _percent_equivalent("1%", 1.0) is False
