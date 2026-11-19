from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


class RejectionDetail(BaseModel):
    code: str | None = None
    description: str | None = None
    quantity: int | None = None


class ProductionRecord(BaseModel):
    """One filled hourly line, shift total, or report-level entry."""

    record_type: Literal["hourly", "shift_total", "report"] = "hourly"
    shift_number: int | None = None
    hour_number: int | None = None
    period: str | None = None
    shot_counter: int | None = None
    total_shots: int | None = None
    ok_shots: int | None = None
    rejected_shots: int | None = None
    produced_parts: int | None = None
    type_of_rejection: str | None = None
    rejection_code: str | None = None
    rejection_reason: str | None = None
    rejection_details: list[RejectionDetail] = Field(default_factory=list)
    type_of_breakdown: str | None = None
    breakdown_code: str | None = None
    breakdown_minutes: int | None = None
    breakdown_reason: str | None = None
    remarks: str | None = None
    missing_fields: list[str] = Field(default_factory=list)
    review_required: list[str] = Field(default_factory=list)


class ShiftMetadata(BaseModel):
    """Shift header fields printed once above each shift's eight hourly rows."""

    shift_number: int
    shift_time_range: str | None = None
    heat_number: str | None = None
    remarks: str | None = None
    operator_signature: str | None = None
    shift_ic_signature: str | None = None


class ProductionReport(BaseModel):
    """Structured extraction from one photographed production report."""

    production_date: str | None = None
    machine_id: str | None = None
    component_name: str | None = None
    number_of_cavities: int | None = None
    target: int | None = None
    target_unit: str | None = None
    operator_name: str | None = None
    shifts: list[ShiftMetadata] = Field(default_factory=list)
    heat_number: str | None = None
    records: list[ProductionRecord] = Field(default_factory=list)
    remarks: str | None = None
    other_entries: list[str] = Field(default_factory=list)
    missing_fields: list[str] = Field(default_factory=list)
    review_required: list[str] = Field(default_factory=list)
