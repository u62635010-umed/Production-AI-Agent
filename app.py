from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import streamlit as st

from src.config import Settings, discover_workbook
from src.excel_manager import DuplicateImageError, ExcelManager, RecordConflictError, SHIFT_TIME_RANGES
from src.production_analysis import calculate_production_summary
from src.schemas import ProductionReport
from src.utils import image_sha256, safe_filename
from src.validation import validate_report
from src.workflow import build_workflow, demo_report


st.set_page_config(page_title="Production AI Agent – HPDC", page_icon="🏭", layout="wide")
st.markdown("""
<style>
  .block-container {padding-top: 1.5rem; max-width: 1440px;}
  .hero {padding: 1.3rem 1.6rem; border-radius: 14px; color: white;
         background: linear-gradient(115deg, #12243a 0%, #1b5b70 100%); margin-bottom: 1rem;}
  .hero h1 {margin: 0; font-size: 2rem;}
  .hero p {margin: .45rem 0 0; color: #d7e9ef;}
  div[data-testid="stMetric"] {background: #f4f7fa; border: 1px solid #e0e7ef; padding: 12px; border-radius: 10px;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="hero"><h1>Production AI Agent – HPDC</h1><p>Review handwritten daily production sheets and update the existing workbook safely.</p></div>', unsafe_allow_html=True)

settings = Settings()
try:
    workbook_path = discover_workbook()
    manager = ExcelManager(workbook_path=workbook_path, settings=settings)
except Exception as exc:
    manager = None
    workbook_path = None
    st.error(f"Excel template is unavailable: {exc}")

if "pending_reports" not in st.session_state:
    st.session_state.pending_reports = {}
if "last_update" not in st.session_state:
    st.session_state.last_update = None


def _is_null(value) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _optional_int(value, label: str):
    value = str(value).strip()
    if not value:
        return None
    try:
        parsed = float(value)
        if not parsed.is_integer():
            raise ValueError
        return int(parsed)
    except ValueError as exc:
        raise ValueError(f"{label} must be a whole number or blank.") from exc


def _rows_to_frame(report: dict) -> pd.DataFrame:
    rows = []
    for record in report.get("records", []):
        rows.append({
            "record_type": record.get("record_type", "hourly"),
            "shift_number": record.get("shift_number"),
            "hour_number": record.get("hour_number"),
            "period": record.get("period"),
            "shot_counter": record.get("shot_counter"),
            "total_shots": record.get("total_shots"),
            "ok_shots": record.get("ok_shots"),
            "rejected_shots": record.get("rejected_shots"),
            "produced_parts": record.get("produced_parts"),
            "type_of_rejection": record.get("type_of_rejection"),
            "rejection_code": record.get("rejection_code"),
            "rejection_reason": record.get("rejection_reason"),
            "rejection_details": json.dumps(record.get("rejection_details", []), ensure_ascii=False),
            "type_of_breakdown": record.get("type_of_breakdown"),
            "breakdown_code": record.get("breakdown_code"),
            "breakdown_minutes": record.get("breakdown_minutes"),
            "breakdown_reason": record.get("breakdown_reason"),
            "remarks": record.get("remarks"),
        })
    columns = ["record_type", "shift_number", "hour_number", "period", "shot_counter", "total_shots", "ok_shots", "rejected_shots",
               "type_of_rejection", "rejection_code", "rejection_reason", "produced_parts", "rejection_details",
               "type_of_breakdown", "breakdown_code", "breakdown_minutes", "breakdown_reason", "remarks"]
    return pd.DataFrame(rows, columns=columns)


def _shift_frame(report: dict) -> pd.DataFrame:
    by_number = {int(shift["shift_number"]): shift for shift in report.get("shifts", [])}
    rows = []
    for number in (1, 2, 3):
        shift = by_number.get(number, {})
        rows.append({
            "shift_number": number,
            "shift_time_range": SHIFT_TIME_RANGES[number],
            "heat_number": shift.get("heat_number"),
            "remarks": shift.get("remarks"),
            "operator_signature": shift.get("operator_signature"),
            "shift_ic_signature": shift.get("shift_ic_signature"),
        })
    return pd.DataFrame(rows)


def _normalize_missing_paths(report: dict) -> list[str]:
    paths = list(report.get("missing_fields", []))
    for index, record in enumerate(report.get("records", [])):
        for field in record.get("missing_fields", []):
            text = str(field).strip()
            if text and not text.startswith("records["):
                text = f"records[{index}].{text}"
            paths.append(text)
    return list(dict.fromkeys(path for path in paths if str(path).strip()))


def _review_flags(report: dict) -> list[str]:
    flags = list(report.get("review_required", []))
    for index, record in enumerate(report.get("records", [])):
        flags.extend(f"records[{index}].{item}" for item in record.get("review_required", []))
    return list(dict.fromkeys(str(flag) for flag in flags if str(flag).strip()))


tabs = st.tabs(["Dashboard", "Upload and Review", "Workbook Status"])

with tabs[0]:
    if manager:
        try:
            summary = calculate_production_summary(manager.history_rows())
            kpis = summary["kpis"]
            cols = st.columns(5)
            cols[0].metric("Total shots", "N/A" if kpis["total_production_shots"] is None else f"{kpis['total_production_shots']:,}")
            cols[1].metric("OK shots", "N/A" if kpis["total_ok_shots"] is None else f"{kpis['total_ok_shots']:,}")
            cols[2].metric("Rejected shots", "N/A" if kpis["total_rejected_shots"] is None else f"{kpis['total_rejected_shots']:,}")
            cols[3].metric("Rejection", "N/A" if kpis["rejection_percentage"] is None else f"{kpis['rejection_percentage']}%")
            cols[4].metric("Breakdown minutes", "N/A" if kpis["breakdown_minutes"] is None else f"{kpis['breakdown_minutes']:,}")
            if summary["daily"]:
                st.subheader("Daily production trend")
                chart_df = pd.DataFrame(summary["daily"]).set_index("production_date")[["total_shots", "ok_shots", "rejected_shots"]]
                st.line_chart(chart_df)
            else:
                st.info("No confirmed production measurements are recorded yet.")
        except Exception as exc:
            st.warning(f"Dashboard summary could not be read: {exc}")
    else:
        st.info("Add the existing Excel template to Templates to enable the dashboard.")

with tabs[1]:
    st.subheader("Upload production sheet photographs")
    files = st.file_uploader("Upload one or more JPG, JPEG, or PNG photographs", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    retain_uploads = st.checkbox("Keep original photographs in data_uploads", value=settings.save_uploads,
                                 help="Photographs may contain production data. Files are retained locally only when this is enabled.")

    upload_meta = []
    if files:
        st.caption(f"{len(files)} photograph(s) selected. Review orientation before extraction.")
        preview_cols = st.columns(min(3, len(files)))
        for index, uploaded in enumerate(files):
            raw = uploaded.getvalue()
            digest = image_sha256(raw)
            rotation = st.selectbox("Rotate image", [0, 90, 180, 270], key=f"rotation_{digest}", format_func=lambda v: f"{v}°")
            with preview_cols[index % len(preview_cols)]:
                st.image(raw, caption=uploaded.name, width="stretch")
            upload_meta.append({"name": uploaded.name, "bytes": raw, "hash": digest, "rotation": rotation})

    action_cols = st.columns([1, 1, 3])
    extract_clicked = action_cols[0].button("Start AI extraction", type="primary", disabled=not files or manager is None, width="stretch")
    demo_clicked = action_cols[1].button("Preview demo data", disabled=manager is None, width="stretch")
    action_cols[2].caption("The demo preview is synthetic and cannot be written to the production workbook.")

    if demo_clicked:
        demo = demo_report()
        st.session_state.pending_reports["demo-preview"] = {
            "report": demo.model_dump(), "filename": "SYNTHETIC DEMO - not a photograph", "hash": "DEMO-NOT-A-REAL-IMAGE",
            "demo": True, "rotation": 0,
        }
        st.rerun()

    if extract_clicked and files and manager:
        progress = st.progress(0, text="Preparing images...")
        graph = build_workflow(manager=manager, settings=settings)
        for index, item in enumerate(upload_meta):
            try:
                if manager.image_already_processed(item["hash"]):
                    st.warning(f"Skipped {item['name']}: this exact photograph is already recorded in the workbook.")
                    progress.progress((index + 1) / len(upload_meta), text=f"Skipped duplicate: {item['name']}")
                    continue
                if retain_uploads:
                    settings.uploads_dir.mkdir(parents=True, exist_ok=True)
                    saved = settings.uploads_dir / f"{item['hash'][:12]}_{safe_filename(item['name'])}"
                    saved.write_bytes(item["bytes"])
                result = graph.invoke({
                    "image_bytes": item["bytes"], "rotation_degrees": item["rotation"],
                    "filename": item["name"], "image_hash": item["hash"], "confirmed": False,
                })
                if result.get("error"):
                    st.error(f"{item['name']}: {result['error']}")
                    continue
                if result.get("report") is None:
                    st.error(f"{item['name']}: no report data was returned.")
                    continue
                st.session_state.pending_reports[item["hash"]] = {
                    "report": result["report"], "filename": item["name"], "hash": item["hash"],
                    "demo": False, "rotation": item["rotation"],
                }
                progress.progress((index + 1) / len(upload_meta), text=f"Extracted {item['name']}; review is required before saving.")
            except Exception as exc:
                st.error(f"{item['name']}: {exc}")
                progress.progress((index + 1) / len(upload_meta), text=f"Could not process {item['name']}")
        progress.empty()
        st.rerun()

    pending = st.session_state.pending_reports
    if pending:
        st.divider()
        labels = {key: f"{item['filename']}  ·  {key[:10]}" for key, item in pending.items()}
        selected = st.selectbox("Select an extracted sheet to review", list(pending), format_func=lambda k: labels[k])
        item = pending[selected]
        report_data = item["report"]
        if item["demo"]:
            st.warning("SYNTHETIC DEMO DATA — this example did not come from a production photograph and cannot be saved to the live workbook.")
        else:
            st.info("Review every extracted field. Unreadable values must be corrected or explicitly confirmed as blank before saving.")

        initial_review_flags = _review_flags(report_data)
        date_needs_review = any("production_date" in flag.lower() or "production date" in flag.lower() for flag in initial_review_flags)
        st.caption("Correct the date, then click Save review edits. The app revalidates it and clears stale date warnings automatically; other genuine issues remain. Once all flags clear, click Confirm and Update Excel.")
        if date_needs_review:
            st.warning("The production date year is uncertain. Check the photograph and enter the full, verified date with a four-digit year. Saving review edits will clear the date warning automatically when the date is valid.")

        with st.form(f"review_{selected}"):
            c1, c2, c3 = st.columns(3)
            production_date = c1.text_input("Production date", value=report_data.get("production_date") or "", help="Enter a verified full date, preferably YYYY-MM-DD. A four-digit year is required.")
            machine_id = c2.text_input("Machine ID", value=report_data.get("machine_id") or "")
            component_name = c3.text_input("Component name", value=report_data.get("component_name") or "")
            c4, c5, c6, c7 = st.columns(4)
            cavities = c4.text_input("Number of cavities", value="" if report_data.get("number_of_cavities") is None else str(report_data["number_of_cavities"]))
            target = c5.text_input("Target", value="" if report_data.get("target") is None else str(report_data["target"]))
            target_unit = c6.text_input("Target unit", value=report_data.get("target_unit") or "")
            operator = c7.text_input("Operator name", value=report_data.get("operator_name") or "")
            st.markdown("**Shift headers (three printed blocks)**")
            st.caption("Heat No., Remarks, and the two signature fields appear once per shift on the paper form.")
            edited_shifts = st.data_editor(
                _shift_frame(report_data),
                key=f"shift_editor_{selected}",
                num_rows="fixed",
                disabled=["shift_number", "shift_time_range"],
                width="stretch",
                column_config={
                    "shift_number": st.column_config.NumberColumn("Shift", format="%d"),
                    "shift_time_range": st.column_config.TextColumn("Printed shift hours"),
                    "heat_number": st.column_config.TextColumn("Heat No."),
                    "remarks": st.column_config.TextColumn("Remarks"),
                    "operator_signature": st.column_config.TextColumn("Opt. Sign (Checked By)"),
                    "shift_ic_signature": st.column_config.TextColumn("Shift I/C Sign (Verified By)"),
                },
            )

            st.markdown("**Hourly rows and shift TOTAL rows**")
            st.caption("Shots, OK shots, rejected shots, and produced parts remain separate. Blank cells marked genuinely missing are stored as numeric 0 with a missing status; values not present in the form stay blank.")
            edited_frame = st.data_editor(
                _rows_to_frame(report_data),
                key=f"records_editor_{selected}",
                num_rows="dynamic",
                width="stretch",
                column_config={
                    "record_type": st.column_config.SelectboxColumn("Record type", options=["hourly", "shift_total", "report"], required=True),
                    "shift_number": st.column_config.NumberColumn("Shift", min_value=1, max_value=3, step=1),
                    "hour_number": st.column_config.NumberColumn("Printed hour row (1–8)", min_value=1, max_value=8, step=1),
                    "shot_counter": st.column_config.NumberColumn("Shot counter", min_value=0, step=1),
                    "total_shots": st.column_config.NumberColumn("Total shots", min_value=0, step=1),
                    "ok_shots": st.column_config.NumberColumn("OK shots", min_value=0, step=1),
                    "rejected_shots": st.column_config.NumberColumn("Rejected shots", min_value=0, step=1),
                    "produced_parts": st.column_config.NumberColumn("Produced parts", min_value=0, step=1),
                    "breakdown_minutes": st.column_config.NumberColumn("Breakdown minutes", min_value=0, step=1),
                    "type_of_rejection": st.column_config.TextColumn("Type of Rej."),
                    "rejection_code": st.column_config.TextColumn("Rejection code"),
                    "rejection_reason": st.column_config.TextColumn("Rejection reason"),
                    "type_of_breakdown": st.column_config.TextColumn("Type of BD"),
                    "breakdown_code": st.column_config.TextColumn("Breakdown code"),
                },
            )
            remarks = st.text_area("Report remarks", value=report_data.get("remarks") or "")
            missing_text = st.text_area(
                "Confirmed genuinely blank fields (one field path per line)",
                value="\n".join(_normalize_missing_paths(report_data)),
                help="Move a field here only after checking the photograph and confirming it is blank. Such numeric fields are written as 0 and excluded from KPI calculations.",
            )
            review_text = st.text_area(
                "Unresolved uncertain fields (correct the data, or move a confirmed blank above)",
                value="\n".join(_review_flags(report_data)),
            )
            submitted = st.form_submit_button("Save review edits", type="primary")

        if submitted:
            try:
                records = []
                for _, row in edited_frame.iterrows():
                    record = {}
                    for field in ("record_type", "period", "type_of_rejection", "rejection_code", "rejection_reason",
                                  "type_of_breakdown", "breakdown_code", "breakdown_reason", "remarks"):
                        value = row.get(field)
                        record[field] = None if _is_null(value) or not str(value).strip() else str(value).strip()
                    if not record.get("record_type"):
                        record["record_type"] = "hourly"
                    for field in ("shift_number", "hour_number", "shot_counter", "total_shots", "ok_shots", "rejected_shots", "produced_parts", "breakdown_minutes"):
                        value = row.get(field)
                        record[field] = None if _is_null(value) else int(value)
                    rejection_json = row.get("rejection_details")
                    if _is_null(rejection_json) or not str(rejection_json).strip():
                        record["rejection_details"] = []
                    else:
                        record["rejection_details"] = json.loads(str(rejection_json))
                    record["missing_fields"] = []
                    record["review_required"] = []
                    records.append(record)
                updated = dict(report_data)
                updated.update({
                    "production_date": production_date.strip() or None,
                    "machine_id": machine_id.strip() or None,
                    "component_name": component_name.strip() or None,
                    "number_of_cavities": _optional_int(cavities, "Number of cavities"),
                    "target": _optional_int(target, "Target"),
                    "target_unit": target_unit.strip() or None,
                    "operator_name": operator.strip() or None,
                    "records": records,
                    "shifts": [
                        {
                            "shift_number": int(row["shift_number"]),
                            "shift_time_range": str(row["shift_time_range"]),
                            "heat_number": None if _is_null(row.get("heat_number")) else str(row.get("heat_number")).strip() or None,
                            "remarks": None if _is_null(row.get("remarks")) else str(row.get("remarks")).strip() or None,
                            "operator_signature": None if _is_null(row.get("operator_signature")) else str(row.get("operator_signature")).strip() or None,
                            "shift_ic_signature": None if _is_null(row.get("shift_ic_signature")) else str(row.get("shift_ic_signature")).strip() or None,
                        }
                        for _, row in edited_shifts.iterrows()
                    ],
                    "remarks": remarks.strip() or None,
                    "missing_fields": [line.strip() for line in missing_text.splitlines() if line.strip()],
                    "review_required": [line.strip() for line in review_text.splitlines() if line.strip()],
                })
                checked = validate_report(ProductionReport.model_validate(updated))
                st.session_state.pending_reports[selected]["report"] = checked.model_dump()
                st.success("Review edits saved. Resolve any remaining flags before updating Excel.")
                st.rerun()
            except Exception as exc:
                st.error(f"Review edits could not be applied: {exc}")

        current_report = validate_report(item["report"])
        st.subheader("Current extraction")
        st.json(current_report.model_dump())
        if current_report.review_required:
            st.error(f"{len(current_report.review_required)} uncertain field(s) still need resolution.")
            for issue in current_report.review_required:
                st.markdown(f"- {issue}")
            if any("production_date" in issue.lower() or "production date" in issue.lower() for issue in current_report.review_required):
                st.info("For a date warning: correct the Production date field above, click Save review edits, and then confirm once this warning clears.")
        elif not item["demo"]:
            st.success("No unresolved validation flags. Confirm the data below to update the workbook.")

        conflicts = {"duplicate_image": False, "conflicts": []}
        if not item["demo"] and current_report.production_date and item["hash"]:
            try:
                conflicts = manager.inspect_conflicts(current_report, item["hash"])
            except Exception as exc:
                st.error(f"Could not check for duplicate records: {exc}")
        if conflicts["duplicate_image"]:
            st.error("This exact photo is already recorded in the workbook. It cannot be added twice.")
        correction_ok = False
        if conflicts["conflicts"]:
            keys = ", ".join(row["record_key"] for row in conflicts["conflicts"])
            st.warning(f"An active record with the same production date, machine, shift, and period exists: {keys}.")
            correction_ok = st.checkbox("This is an intentional correction. Keep the earlier row as Superseded and append the corrected row.", key=f"correction_{selected}")

        confirm = st.button(
            "Confirm and Update Excel",
            type="primary",
            disabled=item["demo"] or bool(current_report.review_required) or conflicts["duplicate_image"] or (bool(conflicts["conflicts"]) and not correction_ok) or manager is None,
            key=f"confirm_{selected}",
        )
        if confirm and manager:
            try:
                graph = build_workflow(manager=manager, settings=settings)
                result = graph.invoke({
                    "report": current_report.model_dump(), "confirmed": True, "filename": item["filename"],
                    "image_hash": item["hash"], "allow_correction": correction_ok, "demo_mode": False,
                })
                if result.get("saved"):
                    st.session_state.last_update = {
                        "records_added": result.get("records_added", 0), "backup": result.get("backup_path"),
                        "filename": item["filename"],
                    }
                    st.success(f"Updated {workbook_path.name} in Templates with {result.get('records_added', 0)} row(s). Backup: {result.get('backup_path')}")
                    del st.session_state.pending_reports[selected]
                    st.rerun()
                elif result.get("error"):
                    st.error(f"Excel update failed: {result['error']}")
                else:
                    st.error("The workflow did not save. Review the data and try again.")
            except (DuplicateImageError, RecordConflictError, PermissionError) as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Excel update failed: {exc}")

with tabs[2]:
    st.subheader("Workbook status")
    if manager:
        try:
            status = manager.status()
            st.write(f"**Selected Excel template:** `{status['workbook']}`")
            st.write(f"**Records currently in Production Data:** {status['records']}")
            st.write(f"**Last successful update:** {status['last_update'] or 'None yet'}")
            st.write(f"**Missing numeric field instances audited:** {status['missing_fields']}")
            st.write(f"**Unresolved errors:** {status['unresolved_errors']}")
            if st.session_state.last_update:
                st.write(f"**Most recent batch rows added:** {st.session_state.last_update['records_added']}")
                st.write(f"**Backup location:** `{st.session_state.last_update['backup']}`")
            st.write("**Backup folder:** `Output/backups`")
            if workbook_path and workbook_path.exists():
                st.download_button("Download current Excel workbook", data=workbook_path.read_bytes(), file_name=workbook_path.name,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.info("Existing sheets Quality and production_summary are preserved. Confirmed records are stored in Production Data in this same workbook. Production Summary, Charts, and Production Audit are updated as part of the confirmation workflow.")
        except Exception as exc:
            st.error(f"Workbook status could not be read: {exc}")
    else:
        st.error("No configured workbook is available. Put the existing .xlsx file in Templates and set its name in config/excel_mapping.yaml.")
