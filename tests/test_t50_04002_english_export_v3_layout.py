from cal_translator.formats.t50_04002.english_export_v3 import laboratory_text_box_spec


def test_laboratory_block_masks_old_label_and_moves_below_logo() -> None:
    spec = laboratory_text_box_spec(
        {
            "information_text_box": {
                "name": "CAL_English_Laboratory_Block",
                "anchor": "B1",
                "width": 360,
                "height": 58,
                "left_offset": -50,
                "top_offset": 8,
                "font_name": "Arial",
                "font_size": 7.5,
                "text_margin_left": 50,
                "text_margin_top": 2,
                "white_fill": True,
                "text": (
                    "Siemens Energy is a trademark licensed by Siemens AG.\n"
                    "Calibration Laboratory\n"
                    "Medellín Highway, km 8.5 - South Side\n"
                    "Tenjo - Cundinamarca"
                ),
                "row_heights": {1: 24, 2: 28},
            }
        }
    )

    assert spec["left_offset"] == -50.0
    assert spec["top_offset"] == 8.0
    assert spec["text_margin_left"] == 50.0
    assert spec["white_fill"] is True
    assert spec["width"] == 360.0
    assert spec["height"] == 58.0
