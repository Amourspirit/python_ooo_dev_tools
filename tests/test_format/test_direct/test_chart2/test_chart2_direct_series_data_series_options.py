from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from com.sun.star.beans import XPropertySet
from com.sun.star.chart2 import XDataSeries
from com.sun.star.chart import XAxisYSupplier
from com.sun.star.chart2 import XDiagramProvider

try:
    from ooodev.office.chart2 import Chart2
    from ooodev.office.chart2 import AxisKind
except ImportError:
    # handle LibreOffice 7.4 bug
    Chart2 = None

from ooodev.format.chart2.direct.series.data_series.options import AlignSeries as SeriesOptAlignSeries
from ooodev.format.chart2.direct.series.data_series.options import Plot as SeriesOptPlot
from ooodev.format.chart2.direct.series.data_series.options import Settings as SeriesOptSettings
from ooodev.format.chart2.direct.series.data_series.options import LegendEntry as SeriesOptLegendEntry
from ooodev.format.chart2.direct.series.data_series.options import MissingValueKind
from ooodev.format.chart2.direct.series.data_series.options import Orientation
from ooodev.utils.data_type.angle import Angle
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

        opt_plot = SeriesOptPlot(chart_doc=chart_doc, missing_values=MissingValueKind.LEAVE_GAP)
        opt_settings = SeriesOptSettings(chart_doc=chart_doc, spacing=111, overlap=22)
        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[opt_plot, opt_settings])

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]
        diagram = chart_doc.getDiagram()
        axis = get_axis_props(diagram, ds1)

        assert diagram.MissingValueTreatment == MissingValueKind.LEAVE_GAP.value
        assert ds1.AttachedAxisIndex == 0
        assert axis.GapWidth == 111
        assert axis.Overlap == 22
        assert diagram.HasYAxis == True
        assert diagram.HasSecondaryYAxis == False

        opt_align = SeriesOptAlignSeries(chart_doc, primary_y_axis=False)
        opt_plot = SeriesOptPlot(chart_doc=chart_doc, missing_values=MissingValueKind.LEAVE_GAP)
        # settings should be last in the styles list
        opt_settings = SeriesOptSettings(chart_doc=chart_doc, spacing=118, overlap=13)
        # orient should be ignored because it is not a valid style for this chart type
        orient = Orientation(chart_doc=chart_doc, clockwise=True, angle=Angle(45))
        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[orient, opt_align, opt_plot, opt_settings])

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]
        diagram = chart_doc.getDiagram()
        axis = get_axis_props(diagram, ds1)

        assert diagram.MissingValueTreatment == MissingValueKind.LEAVE_GAP.value
        assert ds1.AttachedAxisIndex == 1
        assert axis.GapWidth == 118
        assert axis.Overlap == 13
        assert diagram.HasYAxis == False
        assert diagram.HasSecondaryYAxis == True

        opt_align = SeriesOptAlignSeries(chart_doc)
        opt_plot = SeriesOptPlot(
            chart_doc=chart_doc, missing_values=MissingValueKind.USE_ZERO, hidden_cell_values=False
        )
        opt_legned = SeriesOptLegendEntry(chart_doc, hide_legend=False)
        # settings should be last in the styles list
        opt_settings = SeriesOptSettings(chart_doc=chart_doc, spacing=100, overlap=0)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[opt_align, opt_plot, opt_legned, opt_settings])

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]
        diagram = chart_doc.getDiagram()
        axis = get_axis_props(diagram, ds1)

        assert diagram.MissingValueTreatment == MissingValueKind.USE_ZERO.value
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


def test_calc_chart_pie_data_series(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5):
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

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

        range_addr = Calc.get_address(sheet=sheet, range_name="E2:F8")
        chart_doc = Chart2.insert_chart(
            sheet=sheet,
            cells_range=range_addr,
            cell_name="B10",
            width=12,
            height=11,
            diagram_name=ChartTypes.Pie.TEMPLATE_3D.PIE_3D,
        )
        Calc.goto_cell(cell_name="A1", doc=doc)

        Chart2.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="E1"))
        Chart2.set_subtitle(chart_doc=chart_doc, subtitle=Calc.get_string(sheet=sheet, cell_name="F2"))
        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)

        # rotate around horizontal (x-axis) and vertical (y-axis)
        diagram = chart_doc.getFirstDiagram()
        Props.set(
            diagram,
            RotationHorizontal=0,  # -ve rotates bottom edge out of page; default is -60
            RotationVertical=-45,  # -ve rotates left edge out of page; default is 0 (i.e. no rotation)
        )
        orient = Orientation(chart_doc=chart_doc, clockwise=True, angle=Angle(45))
        Chart2.style_data_series(chart_doc=chart_doc, styles=[orient])

        dia = chart_doc.getDiagram()
        supplier = Lo.qi(XAxisYSupplier, dia, True)
        axis_props = supplier.getYAxis()
        clockwise = Props.get(axis_props, "ReverseDirection", None)
        assert clockwise
        assert dia.StartingAngle == 45

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_calc_chart_donut_data_series(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5):
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

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

        range_addr = Calc.get_address(sheet=sheet, range_name="A44:C50")
        chart_doc = Chart2.insert_chart(
            sheet=sheet,
            cells_range=range_addr,
            cell_name="D43",
            width=15,
            height=11,
            diagram_name=ChartTypes.Pie.TEMPLATE_DONUT.DONUT,
        )
        Calc.goto_cell(cell_name="A48", doc=doc)

        Chart2.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A43"))
        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)
        subtitle = f'Outer: {Calc.get_string(sheet=sheet, cell_name="B44")}\nInner: {Calc.get_string(sheet=sheet, cell_name="C44")}'
        Chart2.set_subtitle(chart_doc=chart_doc, subtitle=subtitle)

        orient = Orientation(chart_doc=chart_doc, clockwise=True, angle=Angle(45))
        Chart2.style_data_series(chart_doc=chart_doc, styles=[orient])

        dia = chart_doc.getDiagram()
        supplier = Lo.qi(XAxisYSupplier, dia, True)
        axis_props = supplier.getYAxis()
        clockwise = Props.get(axis_props, "ReverseDirection", None)
        assert clockwise
        assert dia.StartingAngle == 45

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
