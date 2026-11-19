from __future__ import annotations

import json
import math

import pandas as pd


MEASURES = ("total_shots", "ok_shots", "rejected_shots", "produced_parts", "breakdown_minutes")


def _status_map(value) -> dict:
    if isinstance(value, dict):
        return value
    try:
        return json.loads(value or "{}")
    except (TypeError, json.JSONDecodeError):
        return {}


def _numeric(row: dict, field: str):
    statuses = _status_map(row.get("field_status_json"))
    if statuses.get(field) != "confirmed":
        return None
    value = row.get(field)
    if value is None or pd.isna(value):
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def _int_or_none(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    value = float(value)
    return int(value) if value.is_integer() else value


def _measurement_frame(history_rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(history_rows)
    if df.empty:
        return df
    if "status" in df:
        df = df[df["status"].fillna("Active") == "Active"]
    if "demo_mode" in df:
        df = df[~df["demo_mode"].fillna(False).astype(bool)]
    if "record_type" in df:
        hourly_keys = set()
        for _, row in df[df.record_type == "hourly"].iterrows():
            hourly_keys.add((row.get("production_date"), row.get("machine_id"), row.get("shift_number")))
        keep = []
        for _, row in df.iterrows():
            key = (row.get("production_date"), row.get("machine_id"), row.get("shift_number"))
            keep.append(not (row.get("record_type") == "shift_total" and key in hourly_keys))
        df = df.loc[keep]
        df = df[df["record_type"].isin(["hourly", "shift_total"])]
    return df.reset_index(drop=True)


def calculate_production_summary(history_rows: list[dict]) -> dict:
    """Calculate confirmed production metrics; missing-as-zero is excluded."""
    frame = _measurement_frame(history_rows)
    if not frame.empty:
        for field in (*MEASURES, "target"):
            frame[f"_{field}"] = frame.apply(lambda row: _numeric(row.to_dict(), field), axis=1)

    def total(field: str):
        if frame.empty:
            return None
        values = frame[f"_{field}"].dropna()
        return _int_or_none(values.sum()) if not values.empty else None

    total_shots = total("total_shots")
    ok_shots = total("ok_shots")
    rejected = total("rejected_shots")
    breakdown = total("breakdown_minutes")
    rejection_percentage = None
    rejection_complete = bool(not frame.empty and frame["_total_shots"].notna().all() and frame["_rejected_shots"].notna().all())
    if rejection_complete and total_shots not in (None, 0) and rejected is not None:
        rejection_percentage = round(float(rejected) / float(total_shots) * 100, 2)

    daily_map: dict[str, dict] = {}
    machine_map: dict[str, dict] = {}
    shift_map: dict[str, float] = {}
    hourly_map: dict[tuple[str, str], float] = {}
    rejection_map: dict[str, dict[str, float]] = {}
    targets: dict[tuple[str, str], float] = {}
    daily_machines: dict[str, set[str]] = {}
    for _, series in frame.iterrows():
        row = series.to_dict()
        day = str(row.get("production_date") or "")
        machine = str(row.get("machine_id") or "Unknown")
        if not day:
            day = "Unknown"
        daily_machines.setdefault(day, set()).add(machine)
        daily = daily_map.setdefault(day, {"production_date": day, "total_shots": 0, "ok_shots": 0, "rejected_shots": 0,
                                           "_total_shots_seen": False, "_ok_shots_seen": False, "_rejected_shots_seen": False})
        for key in ("total_shots", "ok_shots", "rejected_shots"):
            value = _numeric(row, key)
            if value is not None:
                daily[key] += value
                daily[f"_{key}_seen"] = True
        if _numeric(row, "total_shots") is not None:
            machine_map.setdefault(machine, {"machine_id": machine, "total_shots": 0})["total_shots"] += _numeric(row, "total_shots")
        shift = row.get("shift_number")
        production = _numeric(row, "total_shots")
        if shift is not None and production is not None:
            label = f"Shift {int(shift)}" if float(shift).is_integer() else f"Shift {shift}"
            shift_map[label] = shift_map.get(label, 0) + production
        hour_number = row.get("hour_number")
        hour_label = f"Hour {int(hour_number)}" if hour_number is not None and not pd.isna(hour_number) else row.get("period")
        if row.get("record_type") == "hourly" and hour_label and _numeric(row, "total_shots") is not None:
            label = f"{day} | Shift {row.get('shift_number') or '?'} | {hour_label}"
            hourly_map[(day, label)] = hourly_map.get((day, label), 0) + _numeric(row, "total_shots")
        rej = _numeric(row, "rejected_shots")
        shots = _numeric(row, "total_shots")
        bucket = rejection_map.setdefault(day, {"rejected": 0, "shots": 0, "row_count": 0, "complete_rows": 0})
        bucket["row_count"] += 1
        if shots is not None and rej is not None:
            bucket["complete_rows"] += 1
            bucket["rejected"] += rej
            bucket["shots"] += shots
        if _numeric(row, "target") is not None and str(row.get("target_unit") or "").strip().lower() in {"shot", "shots"}:
            targets.setdefault((day, machine), _numeric(row, "target"))

    daily_rows = []
    rejection_rows = []
    for day, item in sorted(daily_map.items()):
        out = {"production_date": day}
        for key in ("total_shots", "ok_shots", "rejected_shots"):
            out[key] = _int_or_none(item[key]) if item[f"_{key}_seen"] else None
        machines = daily_machines.get(day, set())
        target_value = targets.get((day, next(iter(machines)))) if len(machines) == 1 else None
        if target_value not in (None, 0) and out.get("total_shots") is not None:
            out["target_achievement_pct"] = round(out["total_shots"] / target_value * 100, 2)
        else:
            out["target_achievement_pct"] = None
        daily_rows.append(out)
        bucket = rejection_map.get(day)
        if bucket and bucket["row_count"] == bucket["complete_rows"] and bucket["shots"]:
            rejection_rows.append({"production_date": day, "rejection_percentage": round(bucket["rejected"] / bucket["shots"] * 100, 2)})

    return {
        "kpis": {
            "total_production_shots": total_shots,
            "total_ok_shots": ok_shots,
            "total_rejected_shots": rejected,
            "rejection_percentage": rejection_percentage,
            "breakdown_minutes": breakdown,
            "target_achievement_percentage": next((r["target_achievement_pct"] for r in reversed(daily_rows) if r["target_achievement_pct"] is not None), None),
        },
        "daily": daily_rows,
        "machine": [{"machine_id": k, "total_shots": _int_or_none(v["total_shots"])} for k, v in sorted(machine_map.items())],
        "shift": [{"shift": k, "total_shots": _int_or_none(v)} for k, v in sorted(shift_map.items())],
        "hourly": [{"label": label, "total_shots": _int_or_none(value)} for (_, label), value in sorted(hourly_map.items())],
        "rejection": rejection_rows,
        "good_rejected": [
            {"category": "OK Shots", "count": ok_shots},
            {"category": "Rejected Shots", "count": rejected},
        ] if (ok_shots is not None and rejected is not None and not frame.empty
              and frame["_ok_shots"].notna().all() and frame["_rejected_shots"].notna().all()) else [],
    }


def update_summary_sheet(workbook, summary: dict, sheet_name: str = "Production Summary"):
    marker = "HPDC PRODUCTION SUMMARY - GENERATED BY PRODUCTION AI AGENT"
    if sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        if ws["A1"].value != marker and any(c.value is not None for row in ws.iter_rows() for c in row):
            raise RuntimeError(f"'{sheet_name}' contains existing user data and was not overwritten. Rename it or move it before generating the summary.")
        if ws["A1"].value == marker:
            ws.delete_rows(1, ws.max_row)
    else:
        ws = workbook.create_sheet(sheet_name)
    ws["A1"] = marker
    ws["A1"].font = __import__("openpyxl").styles.Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="1F4E78")
    ws.merge_cells("A1:C1")
    ws["A3"] = "Confirmed Production KPIs"
    ws["A3"].font = __import__("openpyxl").styles.Font(bold=True, size=12)
    ws.append([])
    for col, value in enumerate(("Metric", "Value", "Unit"), 1):
        cell = ws.cell(4, col, value)
        cell.font = __import__("openpyxl").styles.Font(bold=True, color="FFFFFF")
        cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="5B9BD5")
    kpis = summary["kpis"]
    metrics = [
        ("Total Production", kpis["total_production_shots"], "Shots"),
        ("Total OK Production", kpis["total_ok_shots"], "Shots"),
        ("Total Rejected", kpis["total_rejected_shots"], "Shots"),
        ("Rejection Percentage", kpis["rejection_percentage"], "%"),
        ("Breakdown / Downtime", kpis["breakdown_minutes"], "Minutes"),
        ("Target Achievement", kpis["target_achievement_percentage"], "% (target unit must be Shots)"),
    ]
    for rowno, (label, value, unit) in enumerate(metrics, 5):
        ws.cell(rowno, 1, label)
        ws.cell(rowno, 2, "N/A" if value is None else value)
        ws.cell(rowno, 3, unit)

    sections = [
        (1, "Daily Production Data", ["Production Date", "Total Shots", "OK Shots", "Rejected Shots", "Target Achievement %"],
         [{"Production Date": r["production_date"], "Total Shots": r["total_shots"], "OK Shots": r["ok_shots"], "Rejected Shots": r["rejected_shots"], "Target Achievement %": r["target_achievement_pct"]} for r in summary["daily"]]),
        (7, "Shift Production Data", ["Shift", "Total Shots"], [{"Shift": r["shift"], "Total Shots": r["total_shots"]} for r in summary["shift"]]),
        (10, "Hourly Production Data", ["Period", "Total Shots"], [{"Period": r["label"], "Total Shots": r["total_shots"]} for r in summary["hourly"]]),
        (13, "Rejection Percentage Data", ["Production Date", "Rejection Percentage"], [{"Production Date": r["production_date"], "Rejection Percentage": r["rejection_percentage"]} for r in summary["rejection"]]),
        (16, "Good vs Rejected Data", ["Category", "Shots"], [{"Category": r["category"], "Shots": r["count"]} for r in summary["good_rejected"]]),
        (19, "Machine Production Data", ["Machine ID", "Total Shots"], [{"Machine ID": r["machine_id"], "Total Shots": r["total_shots"]} for r in summary["machine"]]),
    ]
    for start_col, title, headers, rows in sections:
        ws.cell(13, start_col, title).font = __import__("openpyxl").styles.Font(bold=True)
        for col, header in enumerate(headers, start_col):
            cell = ws.cell(14, col, header)
            cell.font = __import__("openpyxl").styles.Font(bold=True, color="FFFFFF")
            cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="5B9BD5")
        for rowno, item in enumerate(rows, 15):
            for offset, header in enumerate(headers):
                value = item.get(header)
                ws.cell(rowno, start_col + offset, "N/A" if value is None else value)
    for col, width in {"A": 32, "B": 20, "C": 36}.items():
        ws.column_dimensions[col].width = width
    for col in range(4, 22):
        ws.column_dimensions[ws.cell(1, col).column_letter].width = 22
    ws.freeze_panes = "A4"
    return ws
