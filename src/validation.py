from __future__ import annotations

import re

from dateutil.parser import parse as parse_date

from src.form_constants import SHIFT_TIME_RANGES
from src.schemas import ProductionReport


NUMERIC_REPORT_FIELDS = ("number_of_cavities", "target")
NUMERIC_RECORD_FIELDS = (
    "hour_number", "shot_counter", "total_shots", "ok_shots", "rejected_shots",
    "produced_parts", "breakdown_minutes",
)


def _path_match(items: list[str], path: str, short_name: str) -> bool:
    values = [str(item).strip().lower() for item in items]
    return path.lower() in values or short_name.lower() in values or any(v.startswith(path.lower() + ":") for v in values)


def _normalise_date(value: str | None) -> tuple[str | None, str | None]:
    if not value:
        return None, None
    text = value.strip()
    if not re.search(r"\b(?:19|20)\d{2}\b", text):
        return text, "A full four-digit year is not clear. Confirm the production date."
    try:
        parsed = parse_date(text, dayfirst=True, fuzzy=False)
        return parsed.date().isoformat(), None
    except (ValueError, OverflowError):
        return text, "Production date could not be parsed. Correct it before saving."


def _is_resolved_date_issue(issue: str) -> bool:
    text = str(issue).lower()
    mentions_date = "production_date" in text or "production date" in text
    describes_uncertainty = any(word in text for word in ("year", "ambiguous", "unclear", "not clear", "parse", "confirm"))
    return mentions_date and describes_uncertainty


def validate_report(report: ProductionReport | dict) -> ProductionReport:
    """Normalize dates and printed shift times, then add unresolved issues."""
    report = report if isinstance(report, ProductionReport) else ProductionReport.model_validate(report)
    data = report.model_dump()
    issues = list(data.get("review_required", []))
    missing = list(data.get("missing_fields", []))

    normalized, date_issue = _normalise_date(data.get("production_date"))
    data["production_date"] = normalized
    if date_issue:
        issues.append(f"production_date: {date_issue}")
    for record_index, record in enumerate(data.get("records", [])):
        issues.extend(f"records[{record_index}].{item}" for item in record.get("review_required", []))
    for index, shift in enumerate(data.get("shifts", [])):
        if shift.get("shift_number") not in {1, 2, 3}:
            issues.append(f"shifts[{index}].shift_number: expected Shift 1, Shift 2, or Shift 3.")
        else:
            # These times are printed on the form and should not be inferred from handwriting.
            shift["shift_time_range"] = SHIFT_TIME_RANGES[shift["shift_number"]]

    for field in NUMERIC_REPORT_FIELDS:
        value = data.get(field)
        if value is not None and value < 0:
            issues.append(f"{field}: negative values are not valid; review the extracted value.")

    for index, record in enumerate(data.get("records", [])):
        path = f"records[{index}]"
        for field in NUMERIC_RECORD_FIELDS:
            value = record.get(field)
            full_path = f"{path}.{field}"
            if value is not None and value < 0:
                issues.append(f"{full_path}: negative values are not valid; review the extracted value.")
        shift = record.get("shift_number")
        if shift is not None and shift not in {1, 2, 3}:
            issues.append(f"{path}.shift_number: expected Shift 1, 2, or 3; review the value.")
        if record.get("record_type") in {"hourly", "shift_total"} and shift is None:
            issues.append(f"{path}.shift_number: identify which of the three printed shift blocks contains this row.")
        hour = record.get("hour_number")
        if record.get("record_type") == "hourly" and (hour is None or hour not in range(1, 9)):
            issues.append(f"{path}.hour_number: identify the printed hourly row from 1 through 8.")
        totals = [record.get("total_shots"), record.get("ok_shots"), record.get("rejected_shots")]
        if all(value is not None and value >= 0 for value in totals) and totals[0] != totals[1] + totals[2]:
            issues.append(f"{path}: total_shots does not equal ok_shots + rejected_shots; check units and handwriting.")

    if normalized and date_issue is None:
        # Clear stale date flags after collecting report-level and record-level issues.
        issues = [issue for issue in issues if not _is_resolved_date_issue(issue)]
    data["review_required"] = list(dict.fromkeys(issues))
    data["missing_fields"] = list(dict.fromkeys(missing))
    return ProductionReport.model_validate(data)


def field_status(value, path: str, missing_fields: list[str], review_required: list[str], short_name: str) -> str:
    if _path_match(review_required, path, short_name):
        return "uncertain"
    if _path_match(missing_fields, path, short_name):
        return "missing"
    if value is None:
        return "not_present"
    return "confirmed"
