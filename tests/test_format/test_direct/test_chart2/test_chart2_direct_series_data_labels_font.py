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

from ooodev.format.chart2.direct.series.data_labels.font import FontOnly
from ooodev.format.chart2.direct.series.data_labels.font import FontEffects
from ooodev.format.chart2.direct.series.data_labels.font import FontLine, FontUnderlineEnum

from ooodev.utils.color import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo


def test_calc_chart_data_series_labels_font(loader, copy_fix_calc) -> None:
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

        ft = FontOnly(size=13)

        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[ft])

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]

        assert ds1.CharHeight == 13

        ft = FontOnly(size=16)
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=0, styles=[ft])
        dp = Chart2.get_data_point_props(chart_doc=chart_doc, series_idx=0, idx=0)

        assert dp.CharHeight == 16

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_calc_chart_data_series_labels_font_effects(loader, copy_fix_calc) -> None:
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

        font_effects = FontEffects(
            color=StandardColor.RED_DARK2,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.RED_DARK2),
        )

        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[font_effects])

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]

        assert ds1.CharColor == StandardColor.RED_DARK2
        assert ds1.CharUnderlineColor == StandardColor.RED_DARK2

        font_effects = FontEffects(
            color=StandardColor.PURPLE,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.PURPLE),
        )
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=0, styles=[font_effects])
        dp = Chart2.get_data_point_props(chart_doc=chart_doc, series_idx=0, idx=0)

        assert dp.CharColor == StandardColor.PURPLE
        assert dp.CharUnderlineColor == StandardColor.PURPLE

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
