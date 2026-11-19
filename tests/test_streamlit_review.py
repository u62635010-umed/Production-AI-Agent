from pathlib import Path

from streamlit.testing.v1 import AppTest

from src.excel_manager import ExcelManager
from src.schemas import ProductionReport


def test_review_edit_clears_date_warning_and_enables_confirmation():
    rows_before = ExcelManager().history_rows()
    app = AppTest.from_file(Path(__file__).parents[1] / "app.py", default_timeout=20)
    app.session_state["pending_reports"] = {
        "review-date-test": {
            "report": ProductionReport.model_validate({
                "production_date": "16/11/26",
                "machine_id": "REVIEW-TEST",
                "number_of_cavities": 2,
                "target": 600,
                "shifts": [{"shift_number": 1, "shift_time_range": "07:00 AM to 03:30 PM"}],
                "records": [{
                    "record_type": "hourly",
                    "shift_number": 1,
                    "hour_number": 1,
                    "shot_counter": 606695,
                    "total_shots": None,
                    "ok_shots": None,
                    "rejected_shots": None,
                    "missing_fields": [
                        "records[0].total_shots",
                        "records[0].ok_shots",
                        "records[0].rejected_shots",
                    ],
                }],
                "review_required": [
                    "production_date has an ambiguous year format (16/11/26)",
                    "production_date: A full four-digit year is not clear. Confirm the production date.",
                ],
            }).model_dump(),
            "filename": "review-date-test.jpg",
            "hash": "f" * 64,
            "demo": False,
            "rotation": 0,
        }
    }
    app.run()

    app.text_input[0].set_value("2026-11-16")
    app.text_area[0].set_value("")  # Report remarks
    app.text_area[1].set_value("")  # Confirmed genuinely blank fields
    app.text_area[2].set_value("")  # Unresolved uncertain fields
    app.form_submit_button[0].click()
    app.run()

    saved_review = app.session_state["pending_reports"]["review-date-test"]["report"]
    confirm_button = next(button for button in app.button if button.label == "Confirm and Update Excel")
    assert saved_review["production_date"] == "2026-11-16"
    assert saved_review["review_required"] == []
    assert confirm_button.disabled is False
    assert not app.exception
    assert ExcelManager().history_rows() == rows_before
