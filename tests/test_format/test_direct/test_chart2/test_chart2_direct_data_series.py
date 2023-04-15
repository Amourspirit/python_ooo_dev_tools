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

from ooodev.format.chart2.direct.series.area import Color as SeriesAreaColor
from ooodev.format.chart2.direct.series.area import Gradient as SeriesAreaGradient, PresetGradientKind
from ooodev.format.chart2.direct.series.area import Hatch as SeriesAreaHatch, PresetHatchKind
from ooodev.format.chart2.direct.series.area import Img as SeriesAreaImg, PresetImageKind
from ooodev.format.chart2.direct.series.area import Pattern as SeriesAreaPattern, PresetPatternKind

from ooodev.format.chart2.direct.series.borders import LineProperties as SeriesBorderLineProperties, BorderLineKind

from ooodev.format.chart2.direct.series.transparency import Transparency as SeriesTransparency
from ooodev.format.chart2.direct.series.transparency import (
    Gradient as SeriesGradientTransparency,
    IntensityRange,
    Angle,
)

from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo

from ooodev.utils.color import StandardColor
from ooodev.utils.info import Info


def test_calc_chart_data_series(loader, copy_fix_calc) -> None:
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

        ds = Chart2.get_data_series(chart_doc=chart_doc)
        ds1 = ds[0]
        sc = SeriesAreaColor(StandardColor.GREEN_LIGHT2)

        border = SeriesBorderLineProperties(color=StandardColor.GREEN_DARK2, width=0.5, transparency=15)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[sc, border])
        assert ds1.BorderColor == StandardColor.GREEN_DARK2
        assert ds1.LineTransparence == 15
        assert ds1.BorderTransparency == 15

        # services ('com.sun.star.chart2.DataSeries', 'com.sun.star.chart2.DataPointProperties', 'com.sun.star.beans.PropertySet')
        assert ds1.FillColor == StandardColor.GREEN_LIGHT2

        border = SeriesBorderLineProperties(
            style=BorderLineKind.DASH_DOT_DOT_ROUNDED, color=StandardColor.INDIGO_DARK3, width=0.5, transparency=5
        )
        Chart2.style_data_series(chart_doc=chart_doc, styles=[border])
        assert ds1.BorderColor == StandardColor.INDIGO_DARK3
        assert ds1.LineTransparence == 5
        assert ds1.BorderTransparency == 5

        gradient = SeriesAreaGradient.from_preset(chart_doc=chart_doc, preset=PresetGradientKind.MAHOGANY)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[gradient], idx=0)
        assert ds1.GradientName == str(PresetGradientKind.MAHOGANY)

        hatch = SeriesAreaHatch.from_preset(chart_doc=chart_doc, preset=PresetHatchKind.RED_90_DEGREES_CROSSED)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[hatch], idx=0)
        assert ds1.FillHatchName == str(PresetHatchKind.RED_90_DEGREES_CROSSED)

        pattern = SeriesAreaPattern.from_preset(chart_doc=chart_doc, preset=PresetPatternKind.SHINGLE)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[pattern])
        assert ds1.FillBitmapName == str(PresetPatternKind.SHINGLE)

        image = SeriesAreaImg.from_preset(chart_doc=chart_doc, preset=PresetImageKind.POOL)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[image])
        assert ds1.FillBitmapName == str(PresetImageKind.POOL)

        series_transparency = SeriesTransparency(50)
        Chart2.style_data_series(chart_doc=chart_doc, styles=[series_transparency])
        assert ds1.FillTransparence == 50

        grad_transparent = SeriesGradientTransparency(
            chart_doc=chart_doc, angle=Angle(30), grad_intensity=IntensityRange(0, 100)
        )
        Chart2.style_data_series(chart_doc=chart_doc, styles=[grad_transparent])
        assert ds1.FillTransparenceGradientName.startswith("ChartTransparencyGradient")

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
