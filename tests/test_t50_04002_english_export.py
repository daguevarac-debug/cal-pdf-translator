from cal_translator.formats.t50_04002.english_export import (
    _load_profile,
    build_translation_plan,
    validate_pdf_text,
)


def _payload() -> dict:
    profile = _load_profile()
    dynamic = profile["dynamic_text"]
    return {
        "format_id": "T50-04002",
        "customer": {
            "name": dynamic["customer_name"]["source"],
            "address": dynamic["customer_address"]["source"],
        },
        "calibration_method": dynamic["calibration_method"]["source"],
        "traceability": [
            {"equipment": source}
            for source in dynamic["traceability_equipment"].keys()
        ],
        "notes": list(dynamic["notes"].keys()),
    }


def test_translation_plan_is_controlled_and_has_unique_targets() -> None:
    plan = build_translation_plan(_payload())
    targets = [(operation.sheet, operation.address) for operation in plan]

    assert len(targets) == len(set(targets))
    assert ("Información", "B3") in targets
    assert ("Información", "B26") in targets
    assert ("Información", "B69") in targets
    assert ("Resultados", "B86") in targets
    assert ("Resultados", "B102") in targets


def test_translation_plan_does_not_touch_measurement_data_rows() -> None:
    plan = build_translation_plan(_payload())

    for operation in plan:
        if operation.sheet != "Resultados":
            continue
        row = int("".join(character for character in operation.address if character.isdigit()))
        assert row not in range(10, 34)
        assert row not in range(35, 59)
        assert row not in range(60, 84)


def test_dynamic_translation_uses_reviewed_american_english() -> None:
    plan = build_translation_plan(_payload())
    translated = {operation.logical_name: operation.target for operation in plan}

    assert translated["customer.name"] == "SEDT Distribution Test Field"
    assert "four-wire configuration" in translated["calibration_method"]
    assert translated["traceability[0].equipment"] == "0.1 mΩ Resistance Standard"
    assert translated["notes.86"].startswith("1. These results apply solely")


def test_pdf_text_validation_accepts_complete_english_certificate() -> None:
    profile = _load_profile()
    required = profile["pdf_validation"]["required_phrases"]

    report = validate_pdf_text("\n".join(required), 5, profile)

    assert report["valid"] is True
    assert report["missing_required_phrases"] == []
    assert report["remaining_forbidden_phrases"] == []


def test_pdf_text_validation_rejects_spanish_and_wrong_page_count() -> None:
    profile = _load_profile()
    required = profile["pdf_validation"]["required_phrases"]
    text = "\n".join([*required, "Método de calibración"])

    report = validate_pdf_text(text, 4, profile)

    assert report["valid"] is False
    assert report["page_count"] == 4
    assert "Método de calibración" in report["remaining_forbidden_phrases"]
