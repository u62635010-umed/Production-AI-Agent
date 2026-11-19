import json
from pathlib import Path

import pytest
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

from src.chart_generator import update_charts
from src.config import discover_workbook
from src.excel_manager import DuplicateImageError, ExcelManager, HISTORY_KEYS, RecordConflictError
from src.production_analysis import calculate_production_summary, update_summary_sheet
from src.schemas import ProductionReport
from src.workflow import build_workflow


def sample_report(total=100, image_date="2026-11-16"):
    return ProductionReport.model_validate({
        "production_date": image_date,
        "machine_id": "HPDC-01",
        "component_name": "Bracket A",
        "number_of_cavities": 2,
        "target": 100,
        "target_unit": "shots",
        "shifts": [{"shift_number": 1, "heat_number": "H-101", "remarks": "Checked",
                    "operator_signature": "Op A", "shift_ic_signature": "Lead B"}],
        "records": [{
            "record_type": "hourly", "shift_number": 1, "hour_number": 1,
            "shot_counter": 900, "total_shots": total, "ok_shots": total - 6, "rejected_shots": 6,
            "type_of_rejection": "R1", "rejection_code": "R1", "rejection_reason": "Cold Shot",
            "type_of_breakdown": "B3", "breakdown_code": "B3",
            "produced_parts": 2 * (total - 6), "breakdown_minutes": 12,
        }],
        "review_required": [],
    })


def save_one(manager, report=None, digest="a" * 64, allow_correction=False):
    report = report or sample_report()
    session = manager.begin_session()
    manager.append_report(session, report, digest, "sample.jpg", allow_correction=allow_correction)
    history = manager._history(session.workbook)
    rows = manager.history_rows_from_sheet(history)
    summary = calculate_production_summary(rows)
    update_summary_sheet(session.workbook, summary, manager.summary_sheet)
    update_charts(session.workbook, summary, manager.charts_sheet, manager.summary_sheet)
    return session, summary


def test_correct_cells_and_original_workbook_contents_preserved(workbook_env):
    path, _, manager = workbook_env
    session, _ = save_one(manager)
    backup = manager.commit(session)
    assert backup.exists()
    saved = load_workbook(path, data_only=False)
    assert saved.sheetnames[:2] == ["Quality", "production_summary"]
    assert saved["Quality"]["A1"].value == "=1+1"
    assert saved["Quality"]["B1"].value == "preserve this"
    assert saved["Quality"]["B1"].number_format == "@"
    history = saved["Production Data"]
    headers = {cell.value: cell.column for cell in history[1]}
    assert history.cell(2, headers["Total Shot"]).value == 100
    assert history.cell(2, headers["OK Shot"]).value == 94
    assert history.cell(2, headers["Hour Row (1-8)"]).value == 1
    assert history.cell(2, headers["Shift Heat No."]).value == "H-101"
    assert history.cell(2, headers["Type of Rej."]).value == "R1"
    assert history.cell(2, headers["Type of BD"]).value == "B3"
    assert history.tables["ProductionDataTable"].ref.endswith("AO2")
    assert [column.name for column in history.tables["ProductionDataTable"].tableColumns] == [cell.value for cell in history[1]]
    assert "Production Summary" in saved.sheetnames
    assert "Charts" in saved.sheetnames
    assert "Production Audit" in saved.sheetnames


def test_discovery_falls_back_to_the_only_workbook_if_name_differs(tmp_path):
    templates = tmp_path / "Templates"
    templates.mkdir()
    actual = templates / "actual_company_template.xlsx"
    actual.touch()
    assert discover_workbook({"workbook": "Templates/old_name.xlsx"}, tmp_path) == actual


def test_genuine_missing_numeric_is_written_zero_with_audit_status(workbook_env):
    _, _, manager = workbook_env
    report = ProductionReport.model_validate({
        "production_date": "2026-11-16", "machine_id": "HPDC-02",
        "records": [{"record_type": "hourly", "shift_number": 1, "hour_number": 2, "period": "09:00-10:00",
                     "total_shots": None, "missing_fields": ["records[0].total_shots"]}],
    })
    session = manager.begin_session()
    manager.append_report(session, report, "b" * 64, "missing.jpg")
    ws = manager._history(session.workbook)
    headers = {cell.value: cell.column for cell in ws[1]}
    assert ws.cell(2, headers["Total Shot"]).value == 0
    status = json.loads(ws.cell(2, headers["Field Status (JSON)"]).value)
    assert status["total_shots"] == "missing"


def test_exact_photo_duplicate_is_rejected(workbook_env):
    _, _, manager = workbook_env
    session, _ = save_one(manager)
    manager.commit(session)
    repeated = manager.begin_session()
    with pytest.raises(DuplicateImageError):
        manager.append_report(repeated, sample_report(), "a" * 64, "sample.jpg")


def test_record_conflict_requires_confirmation_and_supersedes_old(workbook_env):
    _, _, manager = workbook_env
    first, _ = save_one(manager, digest="a" * 64)
    manager.commit(first)
    conflict = manager.begin_session()
    with pytest.raises(RecordConflictError):
        manager.append_report(conflict, sample_report(total=101), "b" * 64, "correction.jpg")
    correction, _ = save_one(manager, sample_report(total=101), digest="b" * 64, allow_correction=True)
    manager.commit(correction)
    rows = manager.history_rows()
    assert [r["status"] for r in rows] == ["Superseded", "Active"]


def test_native_charts_are_linked_and_not_duplicated(workbook_env):
    _, _, manager = workbook_env
    session, summary = save_one(manager)
    update_charts(session.workbook, summary, manager.charts_sheet, manager.summary_sheet)
    charts = session.workbook["Charts"]
    assert len(charts._charts) == 5
    manager.commit(session)
    workbook = load_workbook(manager.path)
    charts = workbook["Charts"]
    assert len(charts._charts) == 5
    formulas = [series.val.numRef.f for chart in charts._charts for series in chart.series if series.val and series.val.numRef]
    assert formulas
    assert all("Production Summary" in formula for formula in formulas)


def test_backup_created_before_commit_and_summary_uses_confirmed_data(workbook_env):
    _, _, manager = workbook_env
    session, summary = save_one(manager)
    backup = manager.commit(session)
    assert backup.is_file()
    assert summary["kpis"]["rejection_percentage"] == 6.0


def test_langgraph_confirmation_runs_excel_analysis_and_chart_nodes(workbook_env):
    _, _, manager = workbook_env
    result = build_workflow(manager=manager).invoke({
        "report": sample_report().model_dump(),
        "confirmed": True,
        "filename": "reviewed.jpg",
        "image_hash": "d" * 64,
    })
    assert result["saved"] is True
    assert result["records_added"] == 1
    saved = load_workbook(manager.path)
    assert saved["Production Data"].max_row == 2
    assert len(saved["Charts"]._charts) == 5
    assert Path(result["backup_path"]).is_file()


def test_langgraph_does_not_write_before_user_confirmation(workbook_env):
    path, _, manager = workbook_env
    result = build_workflow(manager=manager).invoke({
        "report": sample_report().model_dump(),
        "confirmed": False,
        "filename": "not_confirmed.jpg",
        "image_hash": "e" * 64,
    })
    assert not result.get("saved", False)
    assert manager.history_rows() == []
    saved = load_workbook(path, data_only=False)
    assert "Production Data" not in saved.sheetnames


def test_blank_template_gets_readable_production_data_structure(workbook_env):
    path, _, manager = workbook_env
    session = manager.begin_session()
    history = manager._history(session.workbook)
    assert history.title == "Production Data"
    assert history.cell(1, 1).value == "Record ID"
    assert "Production Date" in [cell.value for cell in history[1]]
    assert "Machine ID" in [cell.value for cell in history[1]]
    assert "Component Name" in [cell.value for cell in history[1]]
    assert "Total Shot" in [cell.value for cell in history[1]]
    assert "OK Shot" in [cell.value for cell in history[1]]
    assert "Rej. Shot" in [cell.value for cell in history[1]]
    assert "Rejection Reason" in [cell.value for cell in history[1]]
    assert "BD/no. of min" in [cell.value for cell in history[1]]
    assert history.max_row == 1
    manager._audit(session.workbook)
    empty_summary = calculate_production_summary([])
    update_summary_sheet(session.workbook, empty_summary, manager.summary_sheet)
    update_charts(session.workbook, empty_summary, manager.charts_sheet, manager.summary_sheet)
    manager.commit(session)
    saved = load_workbook(path, data_only=False)
    assert {"Production Data", "Production Summary", "Charts"}.issubset(saved.sheetnames)
    assert saved["Production Data"].max_row == 1
    assert len(saved["Charts"]._charts) == 5


def test_shift_times_use_the_printed_form_schedule():
    from src.validation import validate_report

    report = sample_report().model_dump()
    report["shifts"][0]["shift_time_range"] = "incorrect model guess"
    checked = validate_report(report)
    assert checked.shifts[0].shift_time_range == "07:00 AM to 03:30 PM"


def test_legacy_empty_history_sheet_is_migrated_without_losing_original_sheets(workbook_env):
    path, _, manager = workbook_env
    workbook = load_workbook(path)
    legacy = workbook.create_sheet("Production History")
    for column, heading in enumerate(HISTORY_KEYS, 1):
        legacy.cell(1, column, heading)
    table = Table(displayName="ProductionHistoryTable", ref=f"A1:AO1")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    legacy.add_table(table)
    workbook.save(path)
    workbook.close()

    session = manager.begin_session()
    history = manager._history(session.workbook)
    assert history.title == "Production Data"
    assert history.cell(1, 1).value == "Record ID"
    assert "ProductionHistoryTable" not in history.tables
    assert "ProductionDataTable" in history.tables
    assert [column.name for column in history.tables["ProductionDataTable"].tableColumns] == [cell.value for cell in history[1]]
    assert history.max_row == 1
    manager._audit(session.workbook)
    empty_summary = calculate_production_summary([])
    update_summary_sheet(session.workbook, empty_summary, manager.summary_sheet)
    update_charts(session.workbook, empty_summary, manager.charts_sheet, manager.summary_sheet)
    manager.commit(session)

    migrated = load_workbook(path, data_only=False)
    assert "Production History" not in migrated.sheetnames
    assert migrated["Production Data"].max_row == 1
    assert "Quality" in migrated.sheetnames and "production_summary" in migrated.sheetnames

