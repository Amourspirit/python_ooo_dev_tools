from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])


try:
    from ooodev.office.chart2 import Chart2
except ImportError:
    # handle LibreOffice 7.4 bug
    Chart2 = None

from ooodev.format.chart2.direct.axis.label import Orientation as AxisOrientation, DirectionModeKind
from ooodev.format.chart2.direct.axis.label import Show as AxisShow
from ooodev.format.chart2.direct.axis.label import Order as AxisOrder, ChartAxisArrangeOrderType
from ooodev.format.chart2.direct.axis.label import TextFlow as AxisTextFlow

from ooodev.gui.gui import GUI
from ooodev.utils.info import Info
from ooodev.loader.lo import Lo


def test_chart(loader, copy_fix_calc) -> None:
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

        orient_style = AxisOrientation(angle=55, mode=DirectionModeKind.LR_TB)

        Chart2.style_y_axis(chart_doc=chart_doc, styles=[orient_style])
        yaxis = Chart2.get_y_axis(chart_doc=chart_doc)

        assert yaxis.TextRotation == 55
        assert yaxis.WritingMode == DirectionModeKind.LR_TB.value

        orient_style = AxisOrientation(vertical=True)

        Chart2.style_x_axis(chart_doc=chart_doc, styles=[orient_style])
        xaxis = Chart2.get_x_axis(chart_doc=chart_doc)

        assert xaxis.StackCharacters

        show_style = AxisShow(visible=False)
        Chart2.style_x_axis(chart_doc=chart_doc, styles=[show_style])
        assert xaxis.DisplayLabels is False

        order_style = AxisOrder(ChartAxisArrangeOrderType.STAGGER_EVEN)
        Chart2.style_y_axis(chart_doc=chart_doc, styles=[order_style])

        assert yaxis.ArrangeOrder == ChartAxisArrangeOrderType.STAGGER_EVEN

        text_flow_style = AxisTextFlow(True, True)
        Chart2.style_y_axis(chart_doc=chart_doc, styles=[text_flow_style])

        assert yaxis.TextOverlap == True
        assert yaxis.TextBreak == True

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
