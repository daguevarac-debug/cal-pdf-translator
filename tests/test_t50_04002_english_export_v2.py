from cal_translator.formats.t50_04002.english_export_v2 import (
    information_left_header_text,
    merge_controlled_profile,
    optional_result_columns_are_empty,
    select_results_print_layout,
    traceability_presentation,
)


def test_empty_optional_columns_are_excluded_from_pdf_print_area() -> None:
    values = tuple(tuple(None for _ in range(4)) for _ in range(74))

    assert optional_result_columns_are_empty(values) is True
    assert select_results_print_layout(values) == {
        "optional_columns_empty": True,
        "optional_columns_included": False,
        "print_area": "$B$1:$H$102",
        "fit_to_one_page_wide": False,
    }


def test_nonempty_optional_columns_are_preserved_and_fitted_to_width() -> None:
    values = ((None, "PASS", None, None),)

    assert optional_result_columns_are_empty(values) is False
    assert select_results_print_layout(values) == {
        "optional_columns_empty": False,
        "optional_columns_included": True,
        "print_area": "$B$1:$L$102",
        "fit_to_one_page_wide": True,
    }


def test_controlled_profile_adds_residual_spanish_validation() -> None:
    original = {
        "profile_version": 1,
        "pdf_validation": {
            "required_phrases": ["Calibration Certificate"],
            "forbidden_phrases": ["Certificado de Calibración"],
        },
    }
    overrides = {
        "profile_version": 3,
        "pdf_validation": {
            "required_phrases": ["Internal Code", "Calibration Laboratory"],
            "forbidden_phrases": ["Código Interno", "Página"],
        },
    }

    merged = merge_controlled_profile(original, overrides)

    assert merged["profile_version"] == 3
    assert merged["pdf_validation"]["required_phrases"] == [
        "Calibration Certificate",
        "Internal Code",
        "Calibration Laboratory",
    ]
    assert merged["pdf_validation"]["forbidden_phrases"] == [
        "Certificado de Calibración",
        "Código Interno",
        "Página",
    ]
    assert original["profile_version"] == 1


def test_information_header_preserves_logo_and_laboratory_identity() -> None:
    overrides = {
        "information_left_header": (
            "&G\n"
            '&"Arial,Normal"&7Siemens Energy is a trademark licensed by Siemens AG.\n'
            '&"Arial,Normal"&11Calibration Laboratory\n'
            "Medellín Highway, km 8.5 - South Side\n"
            "Tenjo - Cundinamarca"
        )
    }

    header = information_left_header_text(overrides)

    assert header.startswith("&G\n")
    assert "Calibration Laboratory" in header
    assert "Medellín Highway, km 8.5 - South Side" in header
    assert header.endswith("Tenjo - Cundinamarca")


def test_traceability_layout_has_room_for_wrapped_equipment_names() -> None:
    overrides = {
        "traceability_layout": {
            "equipment_font_size": 8.5,
            "row_heights": {
                69: 25,
                70: 25,
                71: 25,
                72: 25,
                73: 30,
            },
        }
    }

    layout = traceability_presentation(overrides)

    assert layout["equipment_font_size"] == 8.5
    assert layout["row_heights"] == {
        69: 25.0,
        70: 25.0,
        71: 25.0,
        72: 25.0,
        73: 30.0,
    }
