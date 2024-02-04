from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

import uno

try:
    from ooodev.office.chart2 import Chart2, AxisKind
except ImportError:
    # handle LibreOffice 7.4 bug
    Chart2 = None

from ooodev.format.chart2.direct.grid import LineProperties as GridLineProperties, BorderLineKind

from ooodev.utils.color import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.loader.lo import Lo


def test_calc_chart_axis_grid(loader, copy_fix_calc) -> None:
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

        grid_style = GridLineProperties(style=BorderLineKind.CONTINUOUS, color=StandardColor.RED, width=0.5)

        props = Chart2.set_grid_lines(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=0, styles=[grid_style])

        assert props.LineColor == StandardColor.RED

        grid_style = GridLineProperties(style=BorderLineKind.DASH_DOT, color=StandardColor.BRICK, width=0.8)
        Chart2.style_grid(chart_doc=chart_doc, axis_val=AxisKind.Y, styles=[grid_style])
        yaxis = Chart2.get_axis(chart_doc=chart_doc, axis_val=AxisKind.Y, idx=0)
        props = yaxis.getGridProperties()
        assert props.LineColor == StandardColor.BRICK

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
