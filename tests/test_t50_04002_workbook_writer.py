from datetime import datetime

import pytest

from cal_translator.formats.t50_04002.workbook_writer import (
    _require_complete_payload,
    build_scalar_write_plan,
    result_display_rows,
)


def _payload() -> dict:
    results = []
    for channel in "ABC":
        for index in range(24):
            results.append(
                {
                    "channel": channel,
                    "range": "1,0 Ω / 45 A" if index < 2 else "10,0 Ω / 25 A",
                    "specified_value": f"{index},000 mΩ",
                    "average_measured_value": f"{index},001 mΩ",
                    "bias": "0,001 mΩ",
                    "maximum_permissible_error": None,
                    "expanded_uncertainty": "0,0058 mΩ",
                    "coverage_factor": "2,0",
                    "cmc": None,
                    "pass_fail": None,
                    "tur": None,
                    "tar": None,
                }
            )

    return {
        "format_id": "T50-04002",
        "certificate_number": "CAL-13078",
        "internal_code": "32040201",
        "equipment": {
            "description": "HIGH CURRENT RESISTANCE METER",
            "manufacturer": "Tettex",
            "serial_number": "151474",
            "model": "2292",
            "customer_identification": "555010000556",
        },
        "customer": {
            "name": "Customer",
            "address": "Street",
            "city": "Tenjo",
            "order_number": "No",
        },
        "calibration_method": "Controlled calibration method.",
        "reception_date": "2026-01-30",
        "calibration_date": "2026-02-02",
        "issue_date": "2026-02-02",
        "environmental_conditions": {
            "maximum_temperature": "22,7 °C",
            "minimum_temperature": "22,2 °C",
            "maximum_relative_humidity": "37 %",
            "minimum_relative_humidity": "37 %",
        },
        "traceability": [
            {
                "equipment": f"Reference {index}",
                "type": "AEG",
                "internal_number": str(40000 + index),
                "calibrated_by": "SET-GAT",
                "certificate_number": f"NU{index:02d}-25",
                "calibration_date": "2025-08-14",
            }
            for index in range(5)
        ],
        "results": results,
        "notes": [f"{index}. Note {index}" for index in range(1, 8)],
        "extraction_warnings": [],
    }


def test_scalar_plan_targets_only_writable_cells() -> None:
    plan = build_scalar_write_plan(_payload())
    addresses = {(operation.sheet, operation.address) for operation in plan}

    assert len(plan) == 22
    assert ("Información", "B5") in addresses
    assert ("Información", "D46") in addresses
    assert ("Información", "L40") in addresses
    assert ("Información", "L47") in addresses

    # Formula and controlled-text cells must never be overwritten.
    assert ("Información", "B6") not in addresses
    assert ("Información", "D43") not in addresses
    assert ("Información", "H44") not in addresses
    assert ("Información", "B3") not in addresses


def test_scalar_plan_converts_dates_and_environmental_values() -> None:
    plan = build_scalar_write_plan(_payload())
    values = {operation.logical_name: operation.value for operation in plan}

    assert values["dates.received"] == datetime(2026, 1, 30)
    assert values["environment.maximum_temperature.measured"] == 22.7
    assert values["environment.minimum_temperature.measured"] == 22.2
    assert values["environment.maximum_humidity.measured"] == 37.0
    assert values["environment.minimum_humidity.correction"] == 0.0


def test_result_layout_uses_three_controlled_24_row_blocks() -> None:
    positioned = result_display_rows(_payload())

    assert len(positioned) == 72
    assert [row for row, _ in positioned[:24]] == list(range(10, 34))
    assert [row for row, _ in positioned[24:48]] == list(range(35, 59))
    assert [row for row, _ in positioned[48:]] == list(range(60, 84))

    # The range label is printed only on the first row of each repeated range.
    assert positioned[0][1]["range"] == "1,0 Ω / 45 A"
    assert positioned[1][1]["range"] == ""
    assert positioned[2][1]["range"] == "10,0 Ω / 25 A"
    assert positioned[24][1]["range"] == "1,0 Ω / 45 A"


def test_complete_payload_requires_exact_controlled_counts() -> None:
    payload = _payload()
    _require_complete_payload(payload)

    payload["results"] = payload["results"][:-1]
    with pytest.raises(RuntimeError, match="72 result rows"):
        _require_complete_payload(payload)
