from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

import uno

try:
    from ooodev.office.chart2 import Chart2
except ImportError:
    # handle LibreOffice 7.4 bug
    Chart2 = None

from ooodev.format.chart2.direct.axis.positioning import AxisLine, ChartAxisPosition
from ooodev.format.chart2.direct.axis.positioning import PositionAxis
from ooodev.format.chart2.direct.axis.positioning import LabelPosition, ChartAxisLabelPosition
from ooodev.format.chart2.direct.axis.positioning import IntervalMarks, MarkKind, ChartAxisMarkPosition

from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo


def test_calc_chart(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5):
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    delay = 0
    from ooodev.office.calc import Calc

    fix_path = cast(Path, copy_fix_calc("col_chart.ods"))

    doc = Calc.open_doc(fix_path)
    try:
        sheet = Calc.get_sheet(doc)
        if not Lo.bridge_connector.headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

        Calc.goto_cell(cell_name="A1", doc=doc)
        chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="col_chart")

        axis_line_style = AxisLine(cross=ChartAxisPosition.VALUE, value=1)

        Chart2.style_x_axis(chart_doc=chart_doc, styles=[axis_line_style])
        xaxis = Chart2.get_x_axis(chart_doc=chart_doc)

        assert xaxis.CrossoverPosition == ChartAxisPosition.VALUE
        assert xaxis.CrossoverValue == 1

        axis_line_style = AxisLine(cross=ChartAxisPosition.ZERO, value=None)
        Chart2.style_x_axis(chart_doc=chart_doc, styles=[axis_line_style])
        assert xaxis.CrossoverPosition == ChartAxisPosition.ZERO
        assert xaxis.CrossoverValue is None

        # does not apply to y axis in this type of chart.
        position_axis_style = PositionAxis(False)
        Chart2.style_x_axis(chart_doc=chart_doc, styles=[position_axis_style])
        sd = xaxis.getScaleData()
        assert sd.ShiftedCategoryPosition == False

        position_axis_style = PositionAxis(True)
        Chart2.style_x_axis(chart_doc=chart_doc, styles=[position_axis_style])
        sd = xaxis.getScaleData()
        assert sd.ShiftedCategoryPosition == True

        label_position_style = LabelPosition(ChartAxisLabelPosition.NEAR_AXIS_OTHER_SIDE)
        Chart2.style_x_axis(chart_doc=chart_doc, styles=[label_position_style])
        assert xaxis.LabelPosition == ChartAxisLabelPosition.NEAR_AXIS_OTHER_SIDE

        label_position_style = LabelPosition()
        Chart2.style_x_axis(chart_doc=chart_doc, styles=[label_position_style])
        assert xaxis.LabelPosition == ChartAxisLabelPosition.NEAR_AXIS

        interval_marks_style = IntervalMarks(major=MarkKind.BOTH, minor=MarkKind.OUTSIDE)
        Chart2.style_x_axis(chart_doc=chart_doc, styles=[interval_marks_style])
        assert xaxis.MajorTickmarks == MarkKind.BOTH.value
        assert xaxis.MinorTickmarks == MarkKind.OUTSIDE.value

        label_position_style = LabelPosition(ChartAxisLabelPosition.OUTSIDE_START)
        # # pos, only valid when Label Position ( set using LabelPosition class) is Outside end or Outside start
        interval_marks_style = IntervalMarks(
            major=MarkKind.OUTSIDE, minor=MarkKind.NONE, pos=ChartAxisMarkPosition.AT_LABELS_AND_AXIS
        )
        # Order of styles is important, label_position_style must be applied first for IntervalMarks.pos to be applied correctly.
        Chart2.style_x_axis(chart_doc=chart_doc, styles=[label_position_style, interval_marks_style])

        assert xaxis.LabelPosition == ChartAxisLabelPosition.OUTSIDE_START
        assert xaxis.MajorTickmarks == MarkKind.OUTSIDE.value
        assert xaxis.MinorTickmarks == MarkKind.NONE.value
        assert xaxis.MarkPosition == ChartAxisMarkPosition.AT_LABELS_AND_AXIS

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
