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

from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.format.chart2.direct.wall.area import (
    PresetGradientKind,
    PresetHatchKind,
    PresetImageKind,
    PresetPatternKind,
)
from ooodev.format.chart2.direct.wall.borders import LineProperties as WallLineProperties
from ooodev.format.chart2.direct.wall.area import Color as WallColor
from ooodev.format.chart2.direct.wall.area import Gradient as WallGradient, IntensityRange
from ooodev.format.chart2.direct.wall.area import Hatch as WallHatch
from ooodev.format.chart2.direct.wall.area import Img as WallImg
from ooodev.format.chart2.direct.wall.area import Pattern as WallPattern
from ooodev.format.chart2.direct.wall.transparency import Transparency as WallTransparency
from ooodev.format.chart2.direct.wall.transparency import Gradient as WallGradientTransparency

from ooodev.utils.color import StandardColor
from ooodev.utils.info import Info
from ooodev.units import Angle


def test_calc_set_styles_wall_chart(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5):
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    delay = 0  # 0 if Lo.bridge_connector.headless else 5_000
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

        chart_color = WallColor(color=StandardColor.GREEN_LIGHT2)
        chart_bdr_line = WallLineProperties(color=StandardColor.GREEN_DARK3, width=0.7)
        Chart2.style_wall(chart_doc=chart_doc, styles=[chart_color, chart_bdr_line])
        wall = chart_doc.getFirstDiagram().getWall()
        assert wall.FillColor == StandardColor.GREEN_LIGHT2

        chart_grad = WallGradient.from_preset(chart_doc, PresetGradientKind.NEON_LIGHT)
        Chart2.style_wall(chart_doc=chart_doc, styles=[chart_grad])
        wall = chart_doc.getFirstDiagram().getWall()
        assert wall.FillGradientName == PresetGradientKind.NEON_LIGHT.value

        chart_hatch = WallHatch.from_preset(chart_doc, PresetHatchKind.GREEN_30_DEGREES)
        Chart2.style_wall(chart_doc=chart_doc, styles=[chart_hatch])
        wall = chart_doc.getFirstDiagram().getWall()
        assert wall.FillHatchName == chart_hatch.prop_hatch_name

        chart_img = WallImg.from_preset(chart_doc, PresetImageKind.ICE_LIGHT)
        Chart2.style_wall(chart_doc=chart_doc, styles=[chart_img])
        wall = chart_doc.getFirstDiagram().getWall()
        assert wall.FillBitmapName == str(PresetImageKind.ICE_LIGHT)

        chart_pattern = WallPattern.from_preset(chart_doc, PresetPatternKind.HORIZONTAL_BRICK)
        Chart2.style_wall(chart_doc=chart_doc, styles=[chart_pattern])

        chart_transparency = WallTransparency(value=50)
        Chart2.style_wall(chart_doc=chart_doc, styles=[chart_color, chart_transparency])
        wall = chart_doc.getFirstDiagram().getWall()
        assert wall.FillTransparence == 50

        # ChartTransparencyGradient 1
        chart_grad_transparent = WallGradientTransparency(
            chart_doc=chart_doc, angle=Angle(30), grad_intensity=IntensityRange(0, 100)
        )
        Chart2.style_wall(chart_doc=chart_doc, styles=[chart_grad_transparent])
        wall = chart_doc.getFirstDiagram().getWall()
        assert wall.FillTransparenceGradientName.startswith("ChartTransparencyGradient")

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
