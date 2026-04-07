"""Chart creation utilities for pptx slide generation.

Creates native Excel-embedded charts via python-pptx's add_chart() API.
"""

from __future__ import annotations

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.util import Pt, Emu


# Map our simple type names to python-pptx chart types
CHART_TYPE_MAP = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
}


def _parse_color(color_str: str) -> RGBColor | None:
    """Parse a hex color string like '#C00000' into RGBColor."""
    if not color_str:
        return None
    hex_str = color_str.lstrip('#')
    if len(hex_str) != 6:
        return None
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))


def parse_width_pct(width_str: str) -> float:
    """Parse a width string like '60%' or '50%' into a float 0.0-1.0."""
    s = width_str.strip().rstrip('%')
    try:
        return float(s) / 100.0
    except ValueError:
        return 0.6


def add_chart_to_slide(slide, chart_el, left, top, width, height):
    """Add a native Excel chart to a slide.

    Args:
        slide: pptx slide object
        chart_el: ChartElement dataclass instance
        left, top, width, height: position and size in EMU
    Returns:
        The chart shape object
    """
    xl_type = CHART_TYPE_MAP.get(chart_el.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    chart_data = CategoryChartData()
    chart_data.categories = chart_el.categories

    for series in chart_el.series:
        chart_data.add_series(series.name, series.values)

    chart_shape = slide.shapes.add_chart(xl_type, left, top, width, height, chart_data)
    chart = chart_shape.chart

    # Title
    if chart_el.title:
        chart.has_title = True
        chart.chart_title.text_frame.text = chart_el.title
        for para in chart.chart_title.text_frame.paragraphs:
            for run in para.runs:
                run.font.size = Pt(14)
                run.font.bold = True
    else:
        chart.has_title = False

    # Legend: auto-show when multiple series, or respect explicit setting
    show_legend = chart_el.legend or len(chart_el.series) > 1
    chart.has_legend = show_legend

    # Plot formatting
    plot = chart.plots[0]

    # Data labels
    if chart_el.labels:
        plot.has_data_labels = True
        data_labels = plot.data_labels
        data_labels.font.size = Pt(11)
        if chart_el.number_format:
            data_labels.number_format = chart_el.number_format
        # Position: outside end for column/bar, above for line
        if chart_el.chart_type in ("column", "bar"):
            data_labels.label_position = XL_LABEL_POSITION.OUTSIDE_END
        elif chart_el.chart_type == "line":
            data_labels.label_position = XL_LABEL_POSITION.ABOVE
        elif chart_el.chart_type == "pie":
            data_labels.label_position = XL_LABEL_POSITION.BEST_FIT

    # Series colors
    for i, series_data in enumerate(chart_el.series):
        color = _parse_color(series_data.color)
        if color:
            series = plot.series[i]
            fill = series.format.fill
            fill.solid()
            fill.fore_color.rgb = color

    return chart_shape
