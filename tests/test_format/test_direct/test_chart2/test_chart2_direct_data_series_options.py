from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from com.sun.star.beans import XPropertySet
from com.sun.star.chart2 import XDataSeries

try:
    from ooodev.office.chart2 import Chart2
    from ooodev.office.chart2 import AxisKind
except ImportError:
    # handle LibreOffice 7.4 bug
    Chart2 = None

from ooodev.format.chart2.direct.series.options import Options as SeriesOptions, MissingValueTreatmentEnum
from ooodev.utils.color import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.kind.chart2_types import ChartTypes
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props


def test_calc_chart_data_series(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5):
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    def get_axis_props(diagram_obj, data_series: XDataSeries) -> XPropertySet:
        is_primary_y = int(Props.get(data_series, "AttachedAxisIndex", 0)) == 0
        if is_primary_y:
            axis = cast(XPropertySet, Props.get(diagram_obj, "YAxis"))
        else:
            axis = cast(XPropertySet, Props.get(diagram_obj, "SecondaryYAxis"))
        return axis

    delay = 0
    from ooodev.office.calc import Calc

    fix_path = cast(Path, copy_fix_calc("chartsData.ods"))

    doc = Calc.open_doc(fix_path)
    try:
        sheet = Calc.get_sheet(doc)
        if not Lo.bridge_connector.headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

        rng_data = Calc.get_range_obj("A2:B8")
        chart_doc = Chart2.insert_chart(
            cells_range=rng_data.get_cell_range_address(), diagram_name=ChartTypes.Column.DEFAULT
        )
        Chart2.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A1"))

        Chart2.set_subtitle(chart_doc=chart_doc, subtitle="Sales by month")

        Chart2.set_x_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A2"))
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B2"))
        Chart2.set_background_colors(
            chart_doc=chart_doc, bg_color=StandardColor.BLUE_LIGHT1, wall_color=StandardColor.BLUE_LIGHT2
        )

        Calc.goto_cell(cell_name="A1", doc=doc)

        opt = SeriesOptions(chart_doc, spacing=111, overlap=22, missing_values=MissingValueTreatmentEnum.LEAVE_GAP)
        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[opt])

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]
        diagram = chart_doc.getDiagram()
        axis = get_axis_props(diagram, ds1)

        assert diagram.MissingValueTreatment == MissingValueTreatmentEnum.LEAVE_GAP.value
        assert ds1.AttachedAxisIndex == 0
        assert axis.GapWidth == 111
        assert axis.Overlap == 22
        assert diagram.HasYAxis == True
        assert diagram.HasSecondaryYAxis == False

        opt = SeriesOptions(
            chart_doc,
            primary_y_axis=False,
            spacing=118,
            overlap=13,
            missing_values=MissingValueTreatmentEnum.LEAVE_GAP,
        )
        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[opt])

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]
        diagram = chart_doc.getDiagram()
        axis = get_axis_props(diagram, ds1)

        assert diagram.MissingValueTreatment == MissingValueTreatmentEnum.LEAVE_GAP.value
        assert ds1.AttachedAxisIndex == 1
        assert axis.GapWidth == 118
        assert axis.Overlap == 13
        assert diagram.HasYAxis == False
        assert diagram.HasSecondaryYAxis == True

        opt = SeriesOptions(
            chart_doc,
            primary_y_axis=True,
            spacing=100,
            overlap=0,
            missing_values=MissingValueTreatmentEnum.USE_ZERO,
            hidden_cell_values=False,
            hide_legend=False,
        )
        Chart2.style_data_series(chart_doc=chart_doc, styles=[opt])

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]
        diagram = chart_doc.getDiagram()
        axis = get_axis_props(diagram, ds1)

        assert diagram.MissingValueTreatment == MissingValueTreatmentEnum.USE_ZERO.value
        assert ds1.AttachedAxisIndex == 0
        assert axis.GapWidth == 100
        assert axis.Overlap == 0
        assert diagram.HasYAxis == True
        assert diagram.HasSecondaryYAxis == False

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
