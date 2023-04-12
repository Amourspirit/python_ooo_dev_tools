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
from ooodev.format.chart2.direct.area import (
    PresetGradientKind,
    PresetHatchKind,
    PresetImageKind,
    PresetPatternKind,
)
from ooodev.format.chart2.direct.borders import LineProperties as ChartLineProperties
from ooodev.format.chart2.direct.area import Color as ChartColor
from ooodev.format.chart2.direct.area import Gradient as ChartGradient
from ooodev.format.chart2.direct.area import Hatch as ChartHatch
from ooodev.format.chart2.direct.area import Img as ChartImg
from ooodev.format.chart2.direct.area import Pattern as ChartPattern
from ooodev.format.chart2.direct.transparency import Transparency as ChartTransparency
from ooodev.format.chart2.direct.transparency import Gradient as ChartGradientTransparency, IntensityRange

from ooodev.utils.color import StandardColor
from ooodev.utils.info import Info
from ooodev.utils.data_type.angle import Angle

if TYPE_CHECKING:
    from com.sun.star.chart2 import Title
    from com.sun.star.style import CharacterProperties


def test_calc_set_styles_chart(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5):
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    delay = 0  # 0 if Lo.bridge_connector.headless else 5_000
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

        chart_color = ChartColor(color=StandardColor.GREEN_LIGHT2)
        chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_color, chart_bdr_line])

        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillColor == StandardColor.GREEN_LIGHT2

        chart_grad = ChartGradient.from_preset(chart_doc, PresetGradientKind.NEON_LIGHT)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad])
        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillGradientName == PresetGradientKind.NEON_LIGHT.value

        chart_hatch = ChartHatch.from_preset(chart_doc, PresetHatchKind.GREEN_30_DEGREES)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_hatch])
        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillHatchName == chart_hatch.prop_hatch_name

        chart_img = ChartImg.from_preset(chart_doc, PresetImageKind.ICE_LIGHT)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_img])
        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillBitmapName == str(PresetImageKind.ICE_LIGHT)

        chart_pattern = ChartPattern.from_preset(chart_doc, PresetPatternKind.HORIZONTAL_BRICK)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_pattern])

        chart_transparency = ChartTransparency(value=50)
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_color, chart_transparency])
        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillTransparence == 50

        # ChartTransparencyGradient 1
        chart_grad_transparent = ChartGradientTransparency(
            chart_doc=chart_doc, angle=Angle(30), grad_intensity=IntensityRange(0, 100)
        )
        Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad_transparent])
        bg_ps = chart_doc.getPageBackground()
        assert bg_ps.FillTransparenceGradientName.startswith("ChartTransparencyGradient")

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
