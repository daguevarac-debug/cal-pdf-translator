import pytest

from cal_translator.formats.t50_04002.english_export_v3 import laboratory_cell_specs


def test_laboratory_identity_uses_two_real_cell_blocks() -> None:
    specs = laboratory_cell_specs(
        {
            "information_cells": [
                {
                    "range": "B1:E1",
                    "row": 1,
                    "row_height": 12,
                    "font_name": "Arial",
                    "font_size": 7,
                    "bold": False,
                    "text": "Siemens Energy is a trademark licensed by Siemens AG.",
                },
                {
                    "range": "B2:E2",
                    "row": 2,
                    "row_height": 34,
                    "font_name": "Arial",
                    "font_size": 9,
                    "bold": False,
                    "text": "Calibration Laboratory\nMedellín Highway, km 8.5 - South Side\nTenjo - Cundinamarca",
                },
            ]
        }
    )

    assert [spec["range"] for spec in specs] == ["B1:E1", "B2:E2"]
    assert specs[0]["row_height"] == 12.0
    assert specs[1]["row_height"] == 34.0
    assert "Calibration Laboratory" in specs[1]["text"]


def test_laboratory_identity_rejects_floating_box_configuration() -> None:
    with pytest.raises(RuntimeError, match="exactly two cell blocks"):
        laboratory_cell_specs({"information_cells": []})
