from __future__ import annotations
import pytest
from typing import cast, TYPE_CHECKING
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

import uno

try:
    from ooodev.office.chart2 import Chart2
except ImportError:
    # handle LibreOffice 7.4 bug
    Chart2 = None

from ooodev.format.chart2.direct.axis.numbers import Numbers, NumberFormatIndexEnum

from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.loader.lo import Lo
from ooodev.utils.kind.zoom_kind import ZoomKind

if TYPE_CHECKING:
    from com.sun.star.chart import ChartAxis
else:
    ChartAxis = object


def test_chart(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5) or Chart2 is None:
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
            Calc.zoom(doc, ZoomKind.ZOOM_100_PERCENT)

        Calc.goto_cell(cell_name="A1", doc=doc)
        chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="col_chart")

        num_style = Numbers(chart_doc, source_format=False, num_format_index=NumberFormatIndexEnum.NUMBER_DEC2)

        Chart2.style_y_axis(chart_doc=chart_doc, styles=[num_style])
        y_axis = cast(ChartAxis, Chart2.get_y_axis(chart_doc=chart_doc))

        assert y_axis.NumberFormat == 2
        assert y_axis.LinkNumberFormatToSource == False

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
