from datetime import datetime

from cal_translator.formats.t50_04002.workbook_writer import build_scalar_write_plan


def _payload() -> dict:
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
        "reception_date": "2026-01-30",
        "calibration_date": "2026-02-02",
        "issue_date": "2026-02-02",
        "environmental_conditions": {
            "maximum_temperature": "22,7 °C",
            "minimum_temperature": "22,2 °C",
            "maximum_relative_humidity": "37 %",
            "minimum_relative_humidity": "37 %",
        },
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
