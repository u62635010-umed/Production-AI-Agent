from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import shutil
import tempfile

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

from src.config import Settings, discover_workbook, load_mapping
from src.form_constants import SHIFT_TIME_RANGES
from src.schemas import ProductionRecord, ProductionReport
from src.validation import NUMERIC_RECORD_FIELDS, NUMERIC_REPORT_FIELDS, field_status, validate_report


HISTORY_HEADER_LABELS = {
    "record_id": "Record ID",
    "record_key": "Duplicate Check Key",
    "status": "Record Status",
    "image_sha256": "Source Image SHA-256",
    "source_file": "Source Photo",
    "confirmed_at": "Confirmed At (UTC)",
    "production_date": "Production Date",
    "machine_id": "Machine ID",
    "component_name": "Component Name",
    "number_of_cavities": "Number of Cavities",
    "target": "Target",
    "target_unit": "Target Unit",
    "operator_name": "Operator Name",
    "record_type": "Row Type",
    "shift_number": "Shift Number",
    "shift_time_range": "Shift Time Range",
    "shift_heat_number": "Shift Heat No.",
    "operator_signature": "Opt. Sign (Checked By)",
    "shift_ic_signature": "Shift I/C Sign (Verified By)",
    "shift_remarks": "Shift Remarks",
    "hour_number": "Hour Row (1-8)",
    "period": "Production Period",
    "shot_counter": "Shot Count",
    "total_shots": "Total Shot",
    "ok_shots": "OK Shot",
    "rejected_shots": "Rej. Shot",
    "type_of_rejection": "Type of Rej.",
    "rejection_code": "Rejection Code",
    "rejection_reason": "Rejection Reason",
    "produced_parts": "Produced Parts",
    "rejection_details_json": "Rejection Details (JSON)",
    "type_of_breakdown": "Type of BD",
    "breakdown_code": "Breakdown Code",
    "breakdown_reason": "Breakdown Reason",
    "breakdown_minutes": "BD/no. of min",
    "record_remarks": "Hourly Remarks",
    "report_remarks": "Report Remarks",
    "other_entries_json": "Other Entries (JSON)",
    "field_status_json": "Field Status (JSON)",
    "missing_fields_json": "Missing Fields (JSON)",
    "demo_mode": "Demo Record",
}
HISTORY_KEYS = list(HISTORY_HEADER_LABELS)
HISTORY_HEADERS = list(HISTORY_HEADER_LABELS.values())
HISTORY_HEADER_TO_KEY = {label: key for key, label in HISTORY_HEADER_LABELS.items()}
LEGACY_HISTORY_HEADERS = HISTORY_KEYS
AUDIT_HEADERS = [
    "image_sha256", "source_file", "event", "record_count", "record_keys_json",
    "missing_fields_json", "recorded_at", "demo_mode",
]
class DuplicateImageError(RuntimeError):
    pass


class RecordConflictError(RuntimeError):
    def __init__(self, conflicts: list[dict]):
        self.conflicts = conflicts
        details = "; ".join(str(item.get("record_key", "")) for item in conflicts)
        super().__init__(f"Matching active production record(s) already exist: {details}")


@dataclass
class WorkbookSession:
    workbook: object
    original_sheetnames: list[str]
    timestamp: str
    records_added: int = 0
    backup_path: Path | None = None
    saved: bool = False


class ExcelManager:
    """Safe, workbook-centric persistence for a configured existing template."""

    def __init__(self, workbook_path: Path | None = None, settings: Settings | None = None):
        self.settings = settings or Settings()
        self.mapping = load_mapping(self.settings.mapping_path)
        self.path = Path(workbook_path) if workbook_path else discover_workbook(self.mapping, self.settings.base_dir)
        if not self.path.exists():
            raise FileNotFoundError(f"Excel workbook not found: {self.path}")

        self.history_sheet = self.mapping.get("history_sheet", "Production Data")
        self.summary_sheet = self.mapping.get("summary_sheet", "Production Summary")
        self.charts_sheet = self.mapping.get("charts_sheet", "Charts")
        self.audit_sheet = self.mapping.get("audit_sheet", "Production Audit")

    def _load(self):
        keep_vba = self.path.suffix.lower() == ".xlsm"
        return load_workbook(self.path, keep_vba=keep_vba, data_only=False)

    def begin_session(self) -> WorkbookSession:
        wb = self._load()
        return WorkbookSession(wb, list(wb.sheetnames), datetime.now(timezone.utc).isoformat(timespec="seconds"))

    def _ensure_table_sheet(self, wb, title: str, headers: list[str], table_name: str):
        ws = wb[title] if title in wb.sheetnames else wb.create_sheet(title)
        current = [ws.cell(1, col).value for col in range(1, len(headers) + 1)] if ws.max_row else []
        populated = any(cell.value is not None for row in ws.iter_rows() for cell in row)
        if populated and current != headers:
            raise RuntimeError(f"Worksheet '{title}' already contains data that does not match the configured schema. It was left unchanged.")
        if not populated:
            for col, name in enumerate(headers, 1):
                cell = ws.cell(1, col, name)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(vertical="center")
            ws.freeze_panes = "A2"
            ws.row_dimensions[1].height = 25
        existing = list(ws.tables.keys())
        if table_name not in existing and ws.max_row >= 1:
            table = Table(displayName=table_name, ref=f"A1:{ws.cell(max(ws.max_row, 1), len(headers)).coordinate}")
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
            ws.add_table(table)
        elif table_name in existing:
            ws.tables[table_name].ref = f"A1:{ws.cell(max(ws.max_row, 1), len(headers)).coordinate}"
        for col, name in enumerate(headers, 1):
            cell = ws.cell(1, col)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
            ws.column_dimensions[cell.column_letter].width = min(max(len(name) + 3, 14), 32)
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 34
        return ws

    def _history(self, wb):
        old_name = "Production History"
        if self.history_sheet not in wb.sheetnames and old_name in wb.sheetnames:
            wb[old_name].title = self.history_sheet
        elif self.history_sheet in wb.sheetnames and old_name in wb.sheetnames:
            raise RuntimeError(
                f"Both '{old_name}' and '{self.history_sheet}' exist. Review and combine them before saving to avoid losing records."
            )

        ws = wb[self.history_sheet] if self.history_sheet in wb.sheetnames else None
        if ws and [ws.cell(1, col).value for col in range(1, len(LEGACY_HISTORY_HEADERS) + 1)] == LEGACY_HISTORY_HEADERS:
            for col, label in enumerate(HISTORY_HEADERS, 1):
                ws.cell(1, col, label)
        if ws and "ProductionHistoryTable" in ws.tables:
            table = ws.tables["ProductionHistoryTable"]
            del ws.tables["ProductionHistoryTable"]
            table.name = "ProductionDataTable"
            table.displayName = "ProductionDataTable"
            ws.add_table(table)
        result = self._ensure_table_sheet(wb, self.history_sheet, HISTORY_HEADERS, "ProductionDataTable")
        table = result.tables["ProductionDataTable"]
        if not table.tableColumns:
            table._initialise_columns()
        if len(table.tableColumns) != len(HISTORY_HEADERS):
            raise RuntimeError("Production Data table column count does not match its configured headers.")
        for column, label in zip(table.tableColumns, HISTORY_HEADERS):
            column.name = label
        result.sheet_view.showGridLines = False
        result.auto_filter.ref = result.tables["ProductionDataTable"].ref
        return result

    @staticmethod
    def history_rows_from_sheet(ws) -> list[dict]:
        return [
            dict(zip(HISTORY_KEYS, row))
            for row in ws.iter_rows(min_row=2, max_col=len(HISTORY_KEYS), values_only=True)
            if any(value is not None for value in row)
        ]

    def _audit(self, wb):
        return self._ensure_table_sheet(wb, self.audit_sheet, AUDIT_HEADERS, "ProductionAuditTable")

    def inspect_conflicts(self, report: ProductionReport | dict, image_hash: str) -> dict:
        report = report if isinstance(report, ProductionReport) else ProductionReport.model_validate(report)
        wb = self._load()
        if self.audit_sheet in wb.sheetnames:
            ws = wb[self.audit_sheet]
            headers = {cell.value: cell.column for cell in ws[1] if cell.value}
            hash_col = headers.get("image_sha256")
            if hash_col:
                if any(ws.cell(row, hash_col).value == image_hash for row in range(2, ws.max_row + 1)):
                    return {"duplicate_image": True, "conflicts": []}
        if self.history_sheet not in wb.sheetnames:
            return {"duplicate_image": False, "conflicts": []}
        history = wb[self.history_sheet]
        key_col = HISTORY_KEYS.index("record_key") + 1
        status_col = HISTORY_KEYS.index("status") + 1
        wanted = [row for row in self._build_rows(report, image_hash, "preview", False) if row.get("record_key")]
        wanted_keys = {row["record_key"] for row in wanted}
        conflicts = []
        for rowno in range(2, history.max_row + 1):
            key = history.cell(rowno, key_col).value
            status = history.cell(rowno, status_col).value if status_col else "Active"
            if key in wanted_keys and status == "Active":
                conflicts.append({"record_key": key, "row": rowno})
        return {"duplicate_image": False, "conflicts": conflicts}

    def image_already_processed(self, image_hash: str) -> bool:
        wb = self._load()
        try:
            if self.audit_sheet not in wb.sheetnames:
                return False
            ws = wb[self.audit_sheet]
            headers = {cell.value: cell.column for cell in ws[1] if cell.value}
            hash_col = headers.get("image_sha256")
            return bool(hash_col and any(ws.cell(row, hash_col).value == image_hash for row in range(2, ws.max_row + 1)))
        finally:
            wb.close()

    def _build_rows(self, report: ProductionReport, image_hash: str, filename: str, demo_mode: bool) -> list[dict]:
        data = report.model_dump()
        report_missing = data.get("missing_fields", [])
        report_review = data.get("review_required", [])
        shift_rows = [
            shift for shift in data.get("shifts", [])
            if any(shift.get(field) for field in ("heat_number", "remarks", "operator_signature", "shift_ic_signature"))
        ]
        source_records = data.get("records") or [
            ProductionRecord(record_type="shift_total", shift_number=shift["shift_number"]).model_dump()
            for shift in shift_rows
        ] or [ProductionRecord(record_type="report").model_dump()]
        shift_lookup = {shift.shift_number: shift.model_dump() for shift in report.shifts}
        rows = []
        for index, record in enumerate(source_records):
            record_missing = list(record.get("missing_fields", [])) + list(report_missing)
            record_review = list(record.get("review_required", [])) + list(report_review)
            shift_metadata = shift_lookup.get(record.get("shift_number"), {})
            prefix = f"records[{index}]"
            field_statuses = {}
            values = {}
            for field in NUMERIC_REPORT_FIELDS:
                status = field_status(data.get(field), field, report_missing, report_review, field)
                field_statuses[field] = status
                value = data.get(field)
                values[field] = 0 if value is None and status == "missing" else value
            for field in NUMERIC_RECORD_FIELDS:
                path = f"{prefix}.{field}"
                status = field_status(record.get(field), path, record_missing, record_review, field)
                field_statuses[field] = status
                value = record.get(field)
                values[field] = 0 if value is None and status == "missing" else value

            row = {
                "record_id": f"{image_hash[:12]}-{index + 1:03d}",
                "record_key": self.record_key(data.get("production_date"), data.get("machine_id"), record),
                "status": "Active",
                "image_sha256": image_hash,
                "source_file": filename,
                "confirmed_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                "production_date": data.get("production_date"),
                "machine_id": data.get("machine_id"),
                "component_name": data.get("component_name"),
                "number_of_cavities": values["number_of_cavities"],
                "target": values["target"],
                "target_unit": data.get("target_unit"),
                "operator_name": data.get("operator_name"),
                "record_type": record.get("record_type", "hourly"),
                "shift_number": record.get("shift_number"),
                "shift_time_range": SHIFT_TIME_RANGES.get(record.get("shift_number")),
                "shift_heat_number": shift_metadata.get("heat_number") or data.get("heat_number"),
                "operator_signature": shift_metadata.get("operator_signature"),
                "shift_ic_signature": shift_metadata.get("shift_ic_signature"),
                "shift_remarks": shift_metadata.get("remarks"),
                "hour_number": values["hour_number"],
                "period": record.get("period"),
                "shot_counter": values["shot_counter"],
                "total_shots": values["total_shots"],
                "ok_shots": values["ok_shots"],
                "rejected_shots": values["rejected_shots"],
                "type_of_rejection": record.get("type_of_rejection"),
                "rejection_code": record.get("rejection_code"),
                "rejection_reason": record.get("rejection_reason"),
                "produced_parts": values["produced_parts"],
                "rejection_details_json": json.dumps([r for r in record.get("rejection_details", [])], ensure_ascii=False),
                "type_of_breakdown": record.get("type_of_breakdown"),
                "breakdown_code": record.get("breakdown_code"),
                "breakdown_minutes": values["breakdown_minutes"],
                "breakdown_reason": record.get("breakdown_reason"),
                "record_remarks": record.get("remarks"),
                "report_remarks": data.get("remarks"),
                "other_entries_json": json.dumps(data.get("other_entries", []), ensure_ascii=False),
                "field_status_json": json.dumps(field_statuses, ensure_ascii=False),
                "missing_fields_json": json.dumps(list(dict.fromkeys(record_missing)), ensure_ascii=False),
                "demo_mode": bool(demo_mode),
            }
            rows.append(row)
        return rows

    @staticmethod
    def record_key(production_date: str | None, machine_id: str | None, record: dict) -> str | None:
        if not production_date or not machine_id:
            return None
        kind = record.get("record_type", "hourly")
        shift = record.get("shift_number")
        if kind == "hourly":
            identity = (f"hour:{record.get('hour_number')}" if record.get("hour_number") is not None else None)
            identity = identity or record.get("period") or (f"counter:{record.get('shot_counter')}" if record.get("shot_counter") is not None else None)
            if shift is None or identity is None:
                return None
            tail = f"shift:{shift}|{identity}"
        elif kind == "shift_total":
            if shift is None:
                return None
            tail = f"shift:{shift}|total"
        else:
            tail = "report"
        return f"{production_date}|{machine_id.strip().lower()}|{tail.lower()}"

    def append_report(
        self,
        session: WorkbookSession,
        report: ProductionReport | dict,
        image_hash: str,
        filename: str,
        allow_correction: bool = False,
        demo_mode: bool = False,
    ) -> int:
        report = report if isinstance(report, ProductionReport) else ProductionReport.model_validate(report)
        report = validate_report(report)
        unresolved = list(report.review_required) + [issue for record in report.records for issue in record.review_required]
        if unresolved:
            raise ValueError("Resolve every uncertain field before updating Excel.")
        rows = self._build_rows(report, image_hash, filename, demo_mode)
        history = self._history(session.workbook)
        audit = self._audit(session.workbook)

        if any(audit.cell(row, 1).value == image_hash for row in range(2, audit.max_row + 1)):
            raise DuplicateImageError("This exact photograph has already been recorded in this workbook.")

        existing_by_key = {}
        key_col = HISTORY_KEYS.index("record_key") + 1
        status_col = HISTORY_KEYS.index("status") + 1
        for rowno in range(2, history.max_row + 1):
            key = history.cell(rowno, key_col).value
            status = history.cell(rowno, status_col).value
            if key and status == "Active":
                existing_by_key.setdefault(key, []).append(rowno)
        wanted_keys = {r["record_key"] for r in rows if r["record_key"]}
        conflicts = sorted((key, rowno) for key in wanted_keys for rowno in existing_by_key.get(key, []))
        if conflicts and not allow_correction:
            raise RecordConflictError([{"record_key": key, "row": rowno} for key, rowno in conflicts])
        for key, rowno in conflicts:
            history.cell(rowno, status_col, "Superseded")

        for row in rows:
            history.append([row.get(key) for key in HISTORY_KEYS])
        history.tables["ProductionDataTable"].ref = f"A1:{history.cell(history.max_row, len(HISTORY_HEADERS)).coordinate}"
        history.auto_filter.ref = history.tables["ProductionDataTable"].ref

        event = "Correction" if conflicts else "Added"
        audit.append([
            image_hash, filename, event, len(rows),
            json.dumps([r["record_key"] for r in rows if r["record_key"]], ensure_ascii=False),
            json.dumps(sorted({field for row in rows for field in json.loads(row["missing_fields_json"])}), ensure_ascii=False),
            session.timestamp, bool(demo_mode),
        ])
        audit.tables["ProductionAuditTable"].ref = f"A1:{audit.cell(audit.max_row, len(AUDIT_HEADERS)).coordinate}"
        session.records_added += len(rows)
        return len(rows)

    def commit(self, session: WorkbookSession) -> Path:
        if session.saved:
            raise RuntimeError("This workbook session has already been saved.")
        self.settings.output_dir.joinpath("backups").mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        backup = self.settings.output_dir / "backups" / f"{self.path.stem}_{stamp}.backup{self.path.suffix}"
        shutil.copy2(self.path, backup)
        fd, temp_name = tempfile.mkstemp(prefix=f".{self.path.stem}_", suffix=self.path.suffix, dir=self.path.parent)
        os.close(fd)
        temp_path = Path(temp_name)
        try:
            session.workbook.save(temp_path)
            keep_vba = self.path.suffix.lower() == ".xlsm"
            verify = load_workbook(temp_path, keep_vba=keep_vba, data_only=False)
            expected_originals = set(session.original_sheetnames)
            if "Production History" in expected_originals and self.history_sheet == "Production Data":
                expected_originals.remove("Production History")
                expected_originals.add("Production Data")
            if not expected_originals.issubset(set(verify.sheetnames)):
                raise RuntimeError("Workbook validation failed: an original worksheet is missing.")
            if self.history_sheet not in verify.sheetnames or self.summary_sheet not in verify.sheetnames or self.charts_sheet not in verify.sheetnames:
                raise RuntimeError("Workbook validation failed: expected history, summary, or chart sheet is missing.")
            verify.close()
            session.workbook.close()
            os.replace(temp_path, self.path)
        except PermissionError as exc:
            raise PermissionError("Excel could not be saved. Close the workbook in Microsoft Excel and try again. A backup was kept in Output/backups.") from exc
        finally:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except OSError:
                    pass
        session.backup_path = backup
        session.saved = True
        return backup

    def history_rows(self) -> list[dict]:
        wb = self._load()
        try:
            if self.history_sheet not in wb.sheetnames:
                return []
            ws = wb[self.history_sheet]
            return self.history_rows_from_sheet(ws)
        finally:
            wb.close()

    def status(self) -> dict:
        rows = self.history_rows()
        return {
            "workbook": self.path,
            "records": len(rows),
            "last_update": max((row.get("confirmed_at") or "" for row in rows), default="") or None,
            "missing_fields": sum(len(json.loads(row.get("missing_fields_json") or "[]")) for row in rows),
            "unresolved_errors": 0,
        }
