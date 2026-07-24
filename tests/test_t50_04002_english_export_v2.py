from cal_translator.formats.t50_04002.english_export_v2 import (
    optional_result_columns_are_empty,
    select_results_print_layout,
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
