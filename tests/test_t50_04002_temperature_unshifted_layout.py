from cal_translator.formats.t50_04002.temperature_english_export_v2 import _load_profile


def test_temperature_information_section_uses_original_template_rows() -> None:
    profile = _load_profile()
    static = {
        (item["sheet"], item["source"]): item["address"]
        for item in profile["static_cells"]
    }
    paragraphs = profile["static_paragraphs"]

    assert static[("Información", "Resultados de calibración")] == "B72"
    assert static[("Información", "Results Calibration")] == "B73"
    assert paragraphs["result_note_1"]["address"] == "B74"
    assert paragraphs["result_note_2"]["address"] == "B76"
