from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass(slots=True)
class EquipmentData:
    description: str | None = None
    manufacturer: str | None = None
    serial_number: str | None = None
    model: str | None = None
    customer_identification: str | None = None


@dataclass(slots=True)
class CustomerData:
    name: str | None = None
    address: str | None = None
    city: str | None = None
    order_number: str | None = None


@dataclass(slots=True)
class EnvironmentalConditions:
    maximum_temperature: str | None = None
    minimum_temperature: str | None = None
    maximum_relative_humidity: str | None = None
    minimum_relative_humidity: str | None = None


@dataclass(slots=True)
class TraceabilityEntry:
    equipment: str
    type: str
    internal_number: str
    calibrated_by: str
    certificate_number: str
    calibration_date: str


@dataclass(slots=True)
class ResultRow:
    channel: str
    range: str | None
    specified_value: str
    average_measured_value: str
    bias: str
    expanded_uncertainty: str
    coverage_factor: str
    maximum_permissible_error: str | None = None
    cmc: str | None = None
    pass_fail: str | None = None
    tur: str | None = None
    tar: str | None = None


@dataclass(slots=True)
class CalibrationCertificate:
    format_id: str
    schema_version: int = 1
    certificate_number: str | None = None
    internal_code: str | None = None
    equipment: EquipmentData = field(default_factory=EquipmentData)
    customer: CustomerData = field(default_factory=CustomerData)
    calibration_method: str | None = None
    environmental_conditions: EnvironmentalConditions = field(default_factory=EnvironmentalConditions)
    reception_date: str | None = None
    calibration_date: str | None = None
    issue_date: str | None = None
    traceability: list[TraceabilityEntry] = field(default_factory=list)
    results: list[ResultRow] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    source_pages: int = 0
    extraction_warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)
