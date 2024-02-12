from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

import uno

# from com.sun.star.lang import XMultiServiceFactory
# from com.sun.star.container import XNameContainer
from ooodev.utils.kind.chart2_types import ChartTypes

try:
    from ooodev.office.chart2 import Chart2
except ImportError:
    Chart2 = None

from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.info import Info
from ooodev.utils.kind.zoom_kind import ZoomKind


def test_insert_chart(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5) or Chart2 is None:
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    from ooodev.utils.color import StandardColor

    delay = 0
    from ooodev.calc import Calc, CalcDoc

    fix_path = cast(Path, copy_fix_calc("chartsData.ods"))

    doc = CalcDoc.open_doc(fix_path)
    try:
        sheet = doc.sheets[0]
        if not Lo.bridge_connector.headless:
            doc.set_visible()
            Lo.delay(500)
            doc.zoom(ZoomKind.ZOOM_100_PERCENT)

        assert len(sheet.charts) == 0
        rng_data = sheet.get_range(range_name="A2:B8").range_obj
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=rng_data.get_cell_range_address(),
            diagram_name=ChartTypes.Column.DEFAULT,
        )
        assert chart_doc is not None

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
        assert len(sheet.charts) == 1
        chart = sheet.charts[0]
        assert chart is not None

        chart_doc = chart.chart_doc
        assert chart_doc is not None
        draw_page = chart.draw_page
        assert draw_page is not None

        chart_title = chart_doc.get_title()
        assert chart_title is None
        chart_title = chart_doc.set_title("Chart Title")
        assert chart_title is not None

        chart_title = chart_doc.get_title()
        assert chart_title is not None

        assert chart_doc.axis_x is not None
        assert chart_doc.axis_y is not None

        x_title = chart_doc.axis_x.get_title()
        assert x_title is None
        y_title = chart_doc.axis_y.get_title()
        assert y_title is None

        x_title = chart_doc.axis_x.set_title("X Axis Title")
        assert x_title is not None
        y_title = chart_doc.axis_y.set_title("Y Axis Title")
        assert y_title is not None

        x_title = chart_doc.axis_x.get_title()
        assert x_title is not None
        y_title = chart_doc.axis_y.get_title()
        assert y_title is not None

        chart_doc.set_bg_color(StandardColor.GRAY_LIGHT2)
        chart_doc.set_wall_color(StandardColor.BLUE_LIGHT3)

    finally:
        doc.close()
