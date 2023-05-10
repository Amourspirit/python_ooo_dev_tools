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

from ooodev.format.chart2.direct.series.data_labels.borders import (
    LineProperties as BorderLineProperties,
    BorderLineKind,
    Intensity,
)
from ooodev.utils.color import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.kind.chart2_types import ChartTypes
from ooodev.utils.lo import Lo


def test_calc_chart_data_series_labels_borders(loader, copy_fix_calc) -> None:
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

        line_props = BorderLineProperties(style=BorderLineKind.CONTINUOUS, color=StandardColor.RED, width=0.5)

        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[line_props])

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]

        assert ds1.LabelBorderColor == StandardColor.RED

        line_props = BorderLineProperties(style=BorderLineKind.DOT, color=StandardColor.GREEN_LIGHT2, width=0.6)
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=0, styles=[line_props])

        dp = Chart2.get_data_point_props(chart_doc=chart_doc, series_idx=0, idx=0)
        assert dp.LabelBorderColor == StandardColor.GREEN_LIGHT2

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
