from __future__ import annotations

from openpyxl.chart import BarChart, LineChart, Reference


OWN_TITLES = {
    "Daily Production (Shots)",
    "Good vs Rejected Shots",
    "Shift-wise Production (Shots)",
    "Hourly Production (Shots)",
    "Rejection Percentage",
}


def _chart_title_text(chart) -> str:
    title = getattr(chart, "title", None)
    try:
        paragraphs = title.tx.rich.p
        return "".join(run.t for para in paragraphs for run in (para.r or []) if run.t)
    except (AttributeError, TypeError):
        pass
    try:
        points = title.tx.strRef.strCache.pt
        return "".join(point.v for point in points if point.v)
    except (AttributeError, TypeError):
        return ""


def _add_line(target_ws, source_ws, data_col: int, category_col: int, title: str, y_title: str, start_row: int, end_row: int, anchor: str, multi_series: bool = False):
    chart = LineChart()
    chart.title = title
    chart.style = 13
    chart.y_axis.title = y_title
    chart.x_axis.title = "Production Date / Period"
    chart.height, chart.width = 9, 17
    if multi_series:
        data = Reference(source_ws, min_col=data_col, max_col=data_col + 2, min_row=start_row, max_row=end_row)
    else:
        data = Reference(source_ws, min_col=data_col, min_row=start_row, max_row=end_row)
    categories = Reference(source_ws, min_col=category_col, min_row=start_row + 1, max_row=end_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    chart.legend.position = "b"
    target_ws.add_chart(chart, anchor)


def _add_bar(target_ws, source_ws, data_col: int, category_col: int, title: str, x_title: str, y_title: str, start_row: int, end_row: int, anchor: str):
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title
    chart.x_axis.title = x_title
    chart.y_axis.title = y_title
    chart.height, chart.width = 9, 17
    data = Reference(source_ws, min_col=data_col, min_row=start_row, max_row=end_row)
    categories = Reference(source_ws, min_col=category_col, min_row=start_row + 1, max_row=end_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    target_ws.add_chart(chart, anchor)


def update_charts(workbook, summary, charts_sheet_name: str = "Charts", summary_sheet_name: str = "Production Summary"):
    if summary_sheet_name not in workbook.sheetnames:
        raise RuntimeError("Production Summary must be created before Charts.")
    charts = workbook[charts_sheet_name] if charts_sheet_name in workbook.sheetnames else workbook.create_sheet(charts_sheet_name)
    source = workbook[summary_sheet_name]
    charts.sheet_view.showGridLines = False
    charts["A30"] = "Charts are linked to confirmed values in Production Summary. Missing values are excluded from analysis."
    charts["A30"].font = __import__("openpyxl").styles.Font(italic=True, color="666666")
    charts._charts = [chart for chart in charts._charts if _chart_title_text(chart) not in OWN_TITLES]

    daily_rows = summary.get("daily", [])
    if daily_rows and any(row.get("total_shots") is not None for row in daily_rows):
        _add_line(charts, source, 2, 1, "Daily Production (Shots)", "Shots", 14, 14 + len(daily_rows), "A1", multi_series=True)
    else:
        _add_line(charts, source, 2, 1, "Daily Production (Shots)", "Shots", 14, 15, "A1", multi_series=True)

    _add_bar(charts, source, 17, 16, "Good vs Rejected Shots", "Category", "Shots", 14, 16, "J1")

    shifts = summary.get("shift", [])
    if shifts:
        _add_bar(charts, source, 8, 7, "Shift-wise Production (Shots)", "Shift", "Total Shots", 14, 14 + len(shifts), "A18")
    else:
        _add_bar(charts, source, 8, 7, "Shift-wise Production (Shots)", "Shift", "Total Shots", 14, 15, "A18")

    hourly = summary.get("hourly", [])
    if hourly:
        _add_line(charts, source, 11, 10, "Hourly Production (Shots)", "Shots", 14, 14 + len(hourly), "J18")
    else:
        _add_line(charts, source, 11, 10, "Hourly Production (Shots)", "Shots", 14, 15, "J18")

    rejection = summary.get("rejection", [])
    if rejection:
        _add_line(charts, source, 14, 13, "Rejection Percentage", "Percent (%)", 14, 14 + len(rejection), "A35")
    else:
        _add_line(charts, source, 14, 13, "Rejection Percentage", "Percent (%)", 14, 15, "A35")

    return charts
