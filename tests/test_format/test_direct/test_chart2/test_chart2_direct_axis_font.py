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

from ooodev.format.chart2.direct.axis.font import FontOnly
from ooodev.format.chart2.direct.axis.font import FontEffects
from ooodev.format.chart2.direct.axis.font import FontLine, FontUnderlineEnum

from ooodev.utils.color import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.loader.lo import Lo


def test_calc_chart_axis_font(loader, copy_fix_calc) -> None:
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

        Chart2.style_x_axis(chart_doc=chart_doc, styles=[ft])
        xaxis = Chart2.get_x_axis(chart_doc=chart_doc)

        assert xaxis.CharHeight == 13

        ft = FontOnly(size=16)
        Chart2.style_y_axis(chart_doc=chart_doc, styles=[ft])

        yaxis = Chart2.get_y_axis(chart_doc=chart_doc)

        yaxis.CharHeight == 16

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_calc_chart_axis_font_effects(loader, copy_fix_calc) -> None:
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

        Chart2.style_x_axis(chart_doc=chart_doc, styles=[font_effects])

        xaxis = Chart2.get_x_axis(chart_doc=chart_doc)

        assert xaxis.CharColor == StandardColor.RED_DARK2
        assert xaxis.CharUnderlineColor == StandardColor.RED_DARK2

        font_effects = FontEffects(
            color=StandardColor.PURPLE,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=StandardColor.PURPLE),
        )
        Chart2.style_y_axis(chart_doc=chart_doc, styles=[font_effects])
        yaxis = Chart2.get_y_axis(chart_doc=chart_doc)

        assert yaxis.CharColor == StandardColor.PURPLE
        assert yaxis.CharUnderlineColor == StandardColor.PURPLE

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
