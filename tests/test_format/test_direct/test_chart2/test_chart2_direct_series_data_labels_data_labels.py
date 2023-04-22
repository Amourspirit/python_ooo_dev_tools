from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooo.dyn.chart2.data_point_label import DataPointLabel

try:
    from ooodev.office.chart2 import Chart2
except ImportError:
    # handle LibreOffice 7.4 bug
    Chart2 = None

from ooodev.format.chart2.direct.series.data_labels.data_labels import NumberFormat
from ooodev.format.chart2.direct.series.data_labels.data_labels import PercentFormat
from ooodev.format.chart2.direct.series.data_labels.data_labels import TextAttribs
from ooodev.format.chart2.direct.series.data_labels.data_labels import AttribOptions
from ooodev.format.chart2.direct.series.data_labels.data_labels import PlacementKind, SeparatorKind
from ooodev.format.chart2.direct.series.data_labels.data_labels import NumberFormatEnum, NumberFormatIndexEnum
from ooodev.format.chart2.direct.series.data_labels.data_labels import Orientation, DirectionModeKind
from ooodev.utils.color import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.kind.chart2_types import ChartTypes
from ooodev.utils.lo import Lo


def test_calc_chart_data_series_labels(loader, copy_fix_calc) -> None:
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

        text_attribs = TextAttribs(show_number=True, show_number_in_percent=True)
        format_percent = PercentFormat(chart_doc=chart_doc)
        attrib_opt = AttribOptions(placement=PlacementKind.INSIDE, separator=SeparatorKind.NEW_LINE)

        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[text_attribs, format_percent, attrib_opt])

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]

        label = cast(DataPointLabel, ds1.Label)
        assert label.ShowNumber is True
        assert label.ShowNumberInPercent is True
        assert label.ShowCategoryName is False
        assert label.ShowLegendSymbol is False
        assert label.ShowCustomLabel is False
        assert label.ShowSeriesName is False

        assert ds1.PercentageNumberFormat is None
        assert ds1.LabelSeparator == "\n"
        assert ds1.LabelPlacement == PlacementKind.INSIDE.value

        text_attribs = TextAttribs(show_number=False, show_number_in_percent=True)
        format_percent = PercentFormat(
            chart_doc=chart_doc, source_format=False, num_format_index=NumberFormatIndexEnum.PERCENT_DEC2
        )
        attrib_opt = AttribOptions(placement=PlacementKind.ABOVE, separator=SeparatorKind.SPACE)
        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[text_attribs, format_percent, attrib_opt])

        label = cast(DataPointLabel, ds1.Label)
        assert label.ShowNumber is False
        assert label.ShowNumberInPercent is True
        assert label.ShowCategoryName is False
        assert label.ShowLegendSymbol is False
        assert label.ShowCustomLabel is False
        assert label.ShowSeriesName is False

        assert ds1.LabelSeparator == " "
        assert ds1.LabelPlacement == PlacementKind.ABOVE.value

        f_format_percent = PercentFormat.from_obj(chart_doc, ds1)
        assert ds1.PercentageNumberFormat == f_format_percent.prop_format_key

        text_attribs = TextAttribs(show_number=True, show_number_in_percent=False)
        format_number = NumberFormat(
            chart_doc=chart_doc, source_format=False, num_format_index=NumberFormatIndexEnum.NUMBER_DEC2
        )
        attrib_opt = AttribOptions(placement=PlacementKind.ABOVE, separator=SeparatorKind.SEMICOLON)
        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[text_attribs, format_number, attrib_opt])

        label = cast(DataPointLabel, ds1.Label)
        assert label.ShowNumber is True
        assert label.ShowNumberInPercent is False
        assert label.ShowCategoryName is False
        assert label.ShowLegendSymbol is False
        assert label.ShowCustomLabel is False
        assert label.ShowSeriesName is False

        assert ds1.LabelSeparator == "; "
        assert ds1.LabelPlacement == PlacementKind.ABOVE.value

        f_format_number = NumberFormat.from_obj(chart_doc, ds1)
        assert ds1.NumberFormat == f_format_number.prop_format_key

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_calc_chart_data_series_labels_rotation(loader, copy_fix_calc) -> None:
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

        rotation = Orientation(angle=60, mode=DirectionModeKind.LR_TB, leaders=True)
        Chart2.style_data_series(chart_doc=chart_doc, idx=0, styles=[rotation])

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]

        assert ds1.TextRotation == 60.0
        assert ds1.WritingMode == DirectionModeKind.LR_TB.value
        assert ds1.ShowCustomLeaderLines

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_calc_chart_data_series_labels_data_point(loader, copy_fix_calc) -> None:
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

        text_attribs = TextAttribs(show_number=True, show_number_in_percent=True)
        format_percent = PercentFormat(chart_doc=chart_doc)
        attrib_opt = AttribOptions(placement=PlacementKind.INSIDE, separator=SeparatorKind.NEW_LINE)

        Chart2.style_data_point(
            chart_doc=chart_doc, series_idx=0, idx=0, styles=[text_attribs, format_percent, attrib_opt]
        )

        dp = Chart2.get_data_point_props(chart_doc=chart_doc, series_idx=0, idx=0)

        label = cast(DataPointLabel, dp.Label)
        assert label.ShowNumber is True
        assert label.ShowNumberInPercent is True
        assert label.ShowCategoryName is False
        assert label.ShowLegendSymbol is False
        assert label.ShowCustomLabel is False
        assert label.ShowSeriesName is False

        assert dp.PercentageNumberFormat is None
        assert dp.LabelSeparator == "\n"
        assert dp.LabelPlacement == PlacementKind.INSIDE.value

        text_attribs = TextAttribs(show_number=False, show_number_in_percent=True)
        format_percent = PercentFormat(
            chart_doc=chart_doc, source_format=False, num_format_index=NumberFormatIndexEnum.PERCENT_DEC2
        )
        attrib_opt = AttribOptions(placement=PlacementKind.ABOVE, separator=SeparatorKind.SPACE)
        Chart2.style_data_point(
            chart_doc=chart_doc, series_idx=0, idx=0, styles=[text_attribs, format_percent, attrib_opt]
        )

        label = cast(DataPointLabel, dp.Label)
        assert label.ShowNumber is False
        assert label.ShowNumberInPercent is True
        assert label.ShowCategoryName is False
        assert label.ShowLegendSymbol is False
        assert label.ShowCustomLabel is False
        assert label.ShowSeriesName is False

        assert dp.LabelSeparator == " "
        assert dp.LabelPlacement == PlacementKind.ABOVE.value

        f_format_percent = PercentFormat.from_obj(chart_doc, dp)
        assert dp.PercentageNumberFormat == f_format_percent.prop_format_key

        text_attribs = TextAttribs(show_number=True, show_number_in_percent=False)
        format_number = NumberFormat(
            chart_doc=chart_doc, source_format=False, num_format_index=NumberFormatIndexEnum.NUMBER_DEC2
        )
        attrib_opt = AttribOptions(placement=PlacementKind.ABOVE, separator=SeparatorKind.SEMICOLON)
        Chart2.style_data_point(
            chart_doc=chart_doc, series_idx=0, idx=0, styles=[text_attribs, format_number, attrib_opt]
        )

        label = cast(DataPointLabel, dp.Label)
        assert label.ShowNumber is True
        assert label.ShowNumberInPercent is False
        assert label.ShowCategoryName is False
        assert label.ShowLegendSymbol is False
        assert label.ShowCustomLabel is False
        assert label.ShowSeriesName is False

        assert dp.LabelSeparator == "; "
        assert dp.LabelPlacement == PlacementKind.ABOVE.value

        f_format_number = NumberFormat.from_obj(chart_doc, dp)
        assert dp.NumberFormat == f_format_number.prop_format_key

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_calc_chart_data_series_labels_rotation_data_point(loader, copy_fix_calc) -> None:
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

        rotation = Orientation(angle=60, mode=DirectionModeKind.LR_TB, leaders=True)
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=0, styles=[rotation])
        dp = Chart2.get_data_point_props(chart_doc=chart_doc, series_idx=0, idx=0)

        assert dp.TextRotation == 60.0
        assert dp.WritingMode == DirectionModeKind.LR_TB.value
        # assert dp.ShowCustomLeaderLines, ShowCustomLeaderLines is not on a data point, only on a data series.

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
