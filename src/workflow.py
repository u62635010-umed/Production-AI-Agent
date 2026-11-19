from __future__ import annotations

from typing import TypedDict, Any

from langgraph.graph import END, StateGraph

from src.chart_generator import update_charts
from src.config import Settings
from src.excel_manager import ExcelManager
from src.image_processor import process_image
from src.production_analysis import calculate_production_summary, update_summary_sheet
from src.schemas import ProductionReport
from src.utils import image_sha256
from src.validation import validate_report
from src.vision_agent import extract_report


class AgentState(TypedDict, total=False):
    image_bytes: bytes | None
    processed_image: bytes | None
    mime_type: str
    rotation_degrees: int
    filename: str
    image_hash: str
    report: dict | None
    review_required: list[str]
    error: str | None
    confirmed: bool
    demo_mode: bool
    allow_correction: bool
    ready_for_excel: bool
    saved: bool
    session: Any
    records_added: int
    backup_path: str | None
    summary: dict


def build_workflow(manager: ExcelManager | None = None, settings: Settings | None = None):
    manager = manager or ExcelManager()
    settings = settings or Settings()

    def image_processing(state: AgentState):
        raw = state.get("image_bytes")
        if raw is None:
            return {}
        try:
            processed, mime = process_image(raw, state.get("rotation_degrees", 0))
            return {"processed_image": processed, "mime_type": mime, "image_hash": state.get("image_hash") or image_sha256(raw)}
        except Exception as exc:
            return {"error": str(exc)}

    def extraction(state: AgentState):
        if state.get("error") or state.get("report") is not None:
            return {}
        if state.get("demo_mode"):
            return {}
        try:
            report = extract_report(state["processed_image"], state.get("mime_type", "image/png"), settings)
            return {"report": report.model_dump()}
        except Exception as exc:
            return {"error": str(exc)}

    def validation(state: AgentState):
        if state.get("error") or state.get("report") is None:
            return {}
        try:
            report = validate_report(state["report"])
            return {"report": report.model_dump(), "review_required": report.review_required}
        except Exception as exc:
            return {"error": f"Extracted data did not match the production schema: {exc}"}

    def human_review(state: AgentState):
        if state.get("error"):
            return {"ready_for_excel": False}
        report = state.get("report") or {}
        unresolved = report.get("review_required", [])
        if unresolved:
            return {"review_required": unresolved, "ready_for_excel": False}
        return {"ready_for_excel": bool(state.get("confirmed"))}

    def route_after_review(state: AgentState):
        return "excel_update" if state.get("ready_for_excel") else END

    def excel_update(state: AgentState):
        try:
            report = ProductionReport.model_validate(state["report"])
            session = manager.begin_session()
            count = manager.append_report(
                session=session,
                report=report,
                image_hash=state.get("image_hash", ""),
                filename=state.get("filename", "uploaded_image"),
                allow_correction=state.get("allow_correction", False),
                demo_mode=state.get("demo_mode", False),
            )
            return {"session": session, "records_added": count}
        except Exception as exc:
            return {"error": str(exc)}

    def production_analysis(state: AgentState):
        if state.get("error") or state.get("session") is None:
            return {}
        try:
            session = state["session"]
            ws = manager._history(session.workbook)
            rows = manager.history_rows_from_sheet(ws)
            summary = calculate_production_summary(rows)
            update_summary_sheet(session.workbook, summary, manager.summary_sheet)
            return {"summary": summary}
        except Exception as exc:
            return {"error": f"Production summary could not be calculated: {exc}"}

    def chart_update(state: AgentState):
        if state.get("error") or state.get("session") is None or state.get("summary") is None:
            return {}
        try:
            session = state["session"]
            update_charts(session.workbook, state["summary"], manager.charts_sheet, manager.summary_sheet)
            backup = manager.commit(session)
            return {"backup_path": str(backup), "saved": True}
        except Exception as exc:
            return {"error": str(exc), "saved": False}

    graph = StateGraph(AgentState)
    graph.add_node("image_processing", image_processing)
    graph.add_node("extraction", extraction)
    graph.add_node("validation", validation)
    graph.add_node("human_review", human_review)
    graph.add_node("excel_update", excel_update)
    graph.add_node("production_analysis", production_analysis)
    graph.add_node("chart_update", chart_update)
    graph.set_entry_point("image_processing")
    graph.add_edge("image_processing", "extraction")
    graph.add_edge("extraction", "validation")
    graph.add_edge("validation", "human_review")
    graph.add_conditional_edges("human_review", route_after_review, {"excel_update": "excel_update", END: END})
    graph.add_edge("excel_update", "production_analysis")
    graph.add_edge("production_analysis", "chart_update")
    graph.add_edge("chart_update", END)
    return graph.compile()


def demo_report() -> ProductionReport:
    """Clearly synthetic preview values. Never returned as photo extraction."""
    return ProductionReport.model_validate({
        "production_date": "2025-01-15",
        "machine_id": "DEMO-MACHINE-01",
        "component_name": "SAMPLE COMPONENT - DEMO ONLY",
        "number_of_cavities": 2,
        "target": 100,
        "target_unit": "shots",
        "operator_name": "DEMO OPERATOR",
        "shifts": [{
            "shift_number": 1,
            "shift_time_range": "07:00 AM to 03:30 PM",
            "heat_number": "DEMO-HEAT",
            "remarks": "DEMO ONLY",
        }],
        "records": [
            {
                "record_type": "hourly", "shift_number": 1, "hour_number": 1,
                "shot_counter": 100, "total_shots": 50, "ok_shots": 47, "rejected_shots": 3,
                "produced_parts": 94, "breakdown_minutes": 0,
                "missing_fields": ["records[0].breakdown_minutes"],
            },
        ],
        "remarks": "SYNTHETIC DEMO DATA. Not extracted from a production photograph.",
        "other_entries": [],
        "missing_fields": [],
        "review_required": [],
    })
