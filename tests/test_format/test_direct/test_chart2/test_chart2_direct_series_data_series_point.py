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

from ooodev.format.chart2.direct.series.data_series.area import Color as SeriesAreaColor
from ooodev.format.chart2.direct.series.data_series.area import Gradient as SeriesAreaGradient, PresetGradientKind
from ooodev.format.chart2.direct.series.data_series.area import Hatch as SeriesAreaHatch, PresetHatchKind
from ooodev.format.chart2.direct.series.data_series.area import Img as SeriesAreaImg, PresetImageKind
from ooodev.format.chart2.direct.series.data_series.area import Pattern as SeriesAreaPattern, PresetPatternKind

from ooodev.format.chart2.direct.series.data_series.borders import (
    LineProperties as SeriesBorderLineProperties,
    BorderLineKind,
)

from ooodev.format.chart2.direct.series.data_series.transparency import Transparency as SeriesTransparency
from ooodev.format.chart2.direct.series.data_series.transparency import (
    Gradient as SeriesGradientTransparency,
    IntensityRange,
    Angle,
)

from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo

from ooodev.utils.color import StandardColor
from ooodev.utils.info import Info


def test_calc_chart_data_series(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5):
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    # BUG: see, Chart2.get_data_points_props(). There is a bug in the API, getDataPointByIndex() is suppose to return XPropertySet,
    # which is does, however, it does not properly implement the XPropertySet interface.
    # setPropertyValue() and getPropertyValue() are not implemented.
    # The Props.set() method can handle this because is has a fallback to set attributes using the setattr() method of python.

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

        sc = SeriesAreaColor(StandardColor.GREEN_LIGHT2)
        # 'com.sun.star.chart2.DataPoint'

        pp = Chart2.get_data_point_props(chart_doc=chart_doc, series_idx=0, idx=-1)
        border = SeriesBorderLineProperties(color=StandardColor.GREEN_DARK2, width=0.5, transparency=15)
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=-1, styles=[sc, border])
        assert pp.BorderColor == StandardColor.GREEN_DARK2
        assert pp.LineTransparence == 15
        assert pp.BorderTransparency == 15
        assert pp.FillColor == StandardColor.GREEN_LIGHT2

        border = SeriesBorderLineProperties(
            style=BorderLineKind.DASH_DOT_DOT_ROUNDED, color=StandardColor.INDIGO_DARK3, width=0.5, transparency=5
        )
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=-1, styles=[border])
        assert pp.BorderColor == StandardColor.INDIGO_DARK3
        assert pp.LineTransparence == 5
        assert pp.BorderTransparency == 5

        gradient = SeriesAreaGradient.from_preset(chart_doc=chart_doc, preset=PresetGradientKind.MAHOGANY)
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=-1, styles=[gradient])
        assert pp.GradientName == str(PresetGradientKind.MAHOGANY)

        hatch = SeriesAreaHatch.from_preset(chart_doc=chart_doc, preset=PresetHatchKind.RED_90_DEGREES_CROSSED)
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=-1, styles=[hatch])
        assert pp.FillHatchName == str(PresetHatchKind.RED_90_DEGREES_CROSSED)

        pattern = SeriesAreaPattern.from_preset(chart_doc=chart_doc, preset=PresetPatternKind.SHINGLE)
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=-1, styles=[pattern])
        assert pp.FillBitmapName == str(PresetPatternKind.SHINGLE)
        image = SeriesAreaImg.from_preset(chart_doc=chart_doc, preset=PresetImageKind.POOL)
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=-1, styles=[image])
        assert pp.FillBitmapName == str(PresetImageKind.POOL)

        series_transparency = SeriesTransparency(50)
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=-1, styles=[series_transparency])
        assert pp.FillTransparence == 50

        grad_transparent = SeriesGradientTransparency(
            chart_doc=chart_doc, angle=Angle(30), grad_intensity=IntensityRange(0, 100)
        )
        Chart2.style_data_point(chart_doc=chart_doc, series_idx=0, idx=-1, styles=[grad_transparent])
        assert pp.FillTransparenceGradientName.startswith("ChartTransparencyGradient")

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
