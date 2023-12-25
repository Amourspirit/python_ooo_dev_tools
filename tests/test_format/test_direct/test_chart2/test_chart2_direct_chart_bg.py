# pylint: disable=wrong-import-order
# pylint: disable=wrong-import-position
# pylint: disable=unused-import
# pylint: disable=useless-import-alias
# pylint: disable=import-outside-toplevel
from __future__ import annotations
import pytest
from typing import cast, TYPE_CHECKING
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
from ooodev.utils.lo import Lo
from ooodev.format.chart2.direct.general.area import (
    PresetGradientKind,
    PresetHatchKind,
    PresetImageKind,
    PresetPatternKind,
)
from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
from ooodev.format.chart2.direct.general.area import Color as ChartColor
from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
from ooodev.format.chart2.direct.general.area import Hatch as ChartHatch
from ooodev.format.chart2.direct.general.area import Img as ChartImg
from ooodev.format.chart2.direct.general.area import Pattern as ChartPattern
from ooodev.format.chart2.direct.general.transparency import Transparency as ChartTransparency
from ooodev.format.chart2.direct.general.transparency import Gradient as ChartGradientTransparency, IntensityRange

from ooodev.utils.color import StandardColor
from ooodev.utils.info import Info
from ooodev.units import Angle

if TYPE_CHECKING:
    from com.sun.star.chart2 import Title
    from com.sun.star.style import CharacterProperties


def test_calc_set_styles_chart(loader, copy_fix_calc) -> None:
    if Chart2 is None or Info.version_info < (7, 5):
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    delay = 0  # 0 if Lo.bridge_connector.headless else 5_000
    from ooodev.office.calc import Calc

    fix_path = cast(Path, copy_fix_calc("col_chart.ods"))

    doc = Calc.open_doc(fix_path)
    try:
        sheet = Calc.get_sheet(doc)
        if not Lo.bridge_connector.headless:  # pylint: disable=no-member
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

        Calc.goto_cell(cell_name="A1", doc=doc)
        chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="col_chart")

        chart_color = ChartColor(color=StandardColor.GREEN_LIGHT2)
        chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_color, chart_bdr_line])

        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillColor == StandardColor.GREEN_LIGHT2  # type: ignore

        chart_grad = ChartGradient.from_preset(chart_doc, PresetGradientKind.NEON_LIGHT)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad])
        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillGradientName == PresetGradientKind.NEON_LIGHT.value  # type: ignore

        chart_hatch = ChartHatch.from_preset(chart_doc, PresetHatchKind.GREEN_30_DEGREES)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_hatch])
        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillHatchName == chart_hatch.prop_hatch_name  # type: ignore

        chart_img = ChartImg.from_preset(chart_doc, PresetImageKind.ICE_LIGHT)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_img])
        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillBitmapName == str(PresetImageKind.ICE_LIGHT)  # type: ignore

        chart_pattern = ChartPattern.from_preset(chart_doc, PresetPatternKind.HORIZONTAL_BRICK)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_pattern])

        chart_transparency = ChartTransparency(value=50)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_color, chart_transparency])
        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillTransparence == 50  # type: ignore

        # ChartTransparencyGradient 1
        chart_grad_transparent = ChartGradientTransparency(
            chart_doc=chart_doc, angle=Angle(30), grad_intensity=IntensityRange(0, 100)
        )
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad_transparent])
        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillTransparenceGradientName.startswith("ChartTransparencyGradient")  # type: ignore

        if not Lo.bridge_connector.headless:  # pylint: disable=no-member
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
