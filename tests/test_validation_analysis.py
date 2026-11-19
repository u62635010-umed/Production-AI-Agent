from src.production_analysis import calculate_production_summary
from src.schemas import ProductionReport
from src.validation import validate_report


def test_unreadable_values_remain_review_required():
    report = ProductionReport.model_validate({
        "production_date": "2026-11-16", "machine_id": "M-10",
        "records": [{"record_type": "hourly", "shift_number": 1, "period": "08:00-09:00",
                     "total_shots": None, "review_required": ["total_shots is unreadable"]}],
    })
    checked = validate_report(report)
    assert checked.review_required
    assert checked.records[0].total_shots is None


def test_corrected_full_date_clears_stale_date_warnings_only():
    report = ProductionReport.model_validate({
        "production_date": "16/11/26",
        "machine_id": "M-10",
        "review_required": [
            "production_date has an ambiguous year format (16/11/26)",
            "production_date: A full four-digit year is not clear. Confirm the production date.",
        ],
        "records": [{
            "record_type": "hourly",
            "shift_number": 1,
            "hour_number": 1,
            "review_required": ["production_date: The year is uncertain."],
        }],
    })
    unresolved = validate_report(report)
    assert len(unresolved.review_required) == 3

    corrected = unresolved.model_dump()
    corrected["production_date"] = "2026-11-16"
    corrected["records"][0]["review_required"] = [
        "production_date: The year is uncertain.",
        "total_shots is unreadable",
    ]
    resolved = validate_report(corrected)
    assert resolved.production_date == "2026-11-16"
    assert resolved.review_required == ["records[0].total_shots is unreadable"]


def test_negative_and_inconsistent_counts_are_flagged():
    report = ProductionReport.model_validate({
        "production_date": "2026-11-16", "machine_id": "M-10",
        "records": [{"record_type": "hourly", "shift_number": 1, "period": "08:00-09:00",
                     "total_shots": -1, "ok_shots": 10, "rejected_shots": 2}],
    })
    issues = validate_report(report).review_required
    assert any("negative" in item for item in issues)
    assert any("total_shots" in item for item in issues)


def test_missing_display_zero_is_excluded_from_kpis():
    rows = [{
        "status": "Active", "demo_mode": False, "record_type": "hourly", "production_date": "2026-11-16",
        "machine_id": "M-10", "shift_number": 1, "period": "08:00-09:00",
        "total_shots": 0, "ok_shots": 0, "rejected_shots": 2, "breakdown_minutes": 0,
        "field_status_json": '{"total_shots":"missing","ok_shots":"confirmed","rejected_shots":"confirmed","breakdown_minutes":"missing"}',
    }]
    result = calculate_production_summary(rows)
    assert result["kpis"]["total_production_shots"] is None
    assert result["kpis"]["total_ok_shots"] == 0
    assert result["kpis"]["total_rejected_shots"] == 2
    assert result["kpis"]["rejection_percentage"] is None


def test_summary_uses_confirmed_values_and_correct_rejection_rate():
    rows = [{
        "status": "Active", "demo_mode": False, "record_type": "hourly", "production_date": "2026-11-16",
        "machine_id": "M-10", "shift_number": 1, "period": "08:00-09:00",
        "total_shots": 100, "ok_shots": 94, "rejected_shots": 6, "breakdown_minutes": 12,
        "field_status_json": '{"total_shots":"confirmed","ok_shots":"confirmed","rejected_shots":"confirmed","breakdown_minutes":"confirmed"}',
    }]
    result = calculate_production_summary(rows)
    assert result["kpis"]["total_production_shots"] == 100
    assert result["kpis"]["total_ok_shots"] == 94
    assert result["kpis"]["total_rejected_shots"] == 6
    assert result["kpis"]["rejection_percentage"] == 6.0
    assert result["kpis"]["breakdown_minutes"] == 12


def test_target_achievement_is_not_combined_across_machines():
    rows = []
    for machine, total, target in (("M-1", 80, 100), ("M-2", 180, 200)):
        rows.append({
            "status": "Active", "demo_mode": False, "record_type": "shift_total", "production_date": "2026-11-16",
            "machine_id": machine, "shift_number": 1, "total_shots": total, "target": target, "target_unit": "shots",
            "field_status_json": '{"total_shots":"confirmed","target":"confirmed"}',
        })
    result = calculate_production_summary(rows)
    assert result["daily"][0]["target_achievement_pct"] is None
    assert result["kpis"]["target_achievement_percentage"] is None
