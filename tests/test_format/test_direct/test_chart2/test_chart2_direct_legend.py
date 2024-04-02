from __future__ import annotations
import pytest
from typing import cast, TYPE_CHECKING
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

import uno

from ooodev.utils.kind.chart2_types import ChartTypes

try:
    from ooodev.office.chart2 import Chart2
except ImportError:
    Chart2 = None

from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.color import StandardColor
from ooodev.utils.info import Info
from ooodev.units import UnitMM100

from ooodev.format.chart2.direct.legend.borders import LineProperties as LegendLineProperties
from ooodev.format.chart2.direct.legend.transparency import (
    Transparency as LegendTransparency,
    Gradient as LegendGradient,
    IntensityRange,
)
from ooodev.format.chart2.direct.legend.area import Color as LegendAreaColor
from ooodev.format.chart2.direct.legend.area import Gradient as LegendAreaGradient, PresetGradientKind
from ooodev.format.chart2.direct.legend.area import Hatch as LegendAreaHatch, PresetHatchKind
from ooodev.format.chart2.direct.legend.area import Pattern as LegendAreaPattern, PresetPatternKind
from ooodev.format.chart2.direct.legend.area import Img as LegendAreaImg, PresetImageKind

from ooodev.format.chart2.direct.legend.font import Font as LegendFont
from ooodev.format.chart2.direct.legend.font import FontEffects as LegendFontEffects, FontStrikeoutEnum
from ooodev.format.chart2.direct.legend.font import FontOnly as LegendFontOnly

from ooodev.format.chart2.direct.general.position_size import Position as ChartShapePosition
from ooodev.format.chart2.direct.general.position_size import Size as ChartShapeSize

from ooodev.format.chart2.direct.legend.position_size import (
    Position as ChartLegendPosition,
    LegendPosition,
    DirectionModeKind,
)

if TYPE_CHECKING:
    from com.sun.star.awt import Size as UnoSize
    from ooo.dyn.awt.point import Point as UnoPoint
    from com.sun.star.chart2 import Legend  # service
else:
    Legend = object


def test_calc_set_styles_legend(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5) or Chart2 is None:
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

        range_addr = Calc.get_address(sheet=sheet, range_name="E2:F8")
        chart_name = "MyChart"
        chart_doc = Chart2.insert_chart(
            sheet=sheet,
            cells_range=range_addr,
            cell_name="B10",
            width=12,
            height=11,
            diagram_name=ChartTypes.Pie.TEMPLATE_3D.PIE_3D,
            chart_name=chart_name,
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
        legend = cast(Legend, Chart2.get_legend(chart_doc=chart_doc))

        legend_line_style = LegendLineProperties(color=StandardColor.MAGENTA, width=0.8, transparency=20)
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_line_style])

        line_width = legend_line_style.prop_width.get_value_mm100()
        assert legend.LineWidth in range(line_width - 2, line_width + 3)  # +/- 2mm100
        assert legend.LineColor == StandardColor.MAGENTA
        assert legend.LineTransparence == 20

        legend_color_style = LegendAreaColor(color=StandardColor.GREEN_LIGHT2)
        legend_bg_transparency_style = LegendTransparency(0)
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_bg_transparency_style, legend_color_style])
        assert legend.FillColor == StandardColor.GREEN_LIGHT2

        legend_area_gradient_style = LegendAreaGradient.from_preset(
            chart_doc=chart_doc, preset=PresetGradientKind.MAHOGANY
        )
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_area_gradient_style])
        assert legend.FillGradientName == str(PresetGradientKind.MAHOGANY)

        legend_hatch_style = LegendAreaHatch.from_preset(chart_doc=chart_doc, preset=PresetHatchKind.GREEN_60_DEGREES)
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_hatch_style])
        assert legend.FillHatchName == str(PresetHatchKind.GREEN_60_DEGREES)

        legend_pattern_style = LegendAreaPattern.from_preset(
            chart_doc=chart_doc, preset=PresetPatternKind.HORIZONTAL_BRICK
        )
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_pattern_style])
        assert legend.FillBitmapName == str(PresetPatternKind.HORIZONTAL_BRICK)

        legend_img_style = LegendAreaImg.from_preset(chart_doc=chart_doc, preset=PresetImageKind.ICE_LIGHT)
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_img_style])
        assert legend.FillBitmapName == str(PresetImageKind.ICE_LIGHT)

        legend_transparency_gradient = LegendGradient(chart_doc, angle=45, grad_intensity=IntensityRange(0, 100))
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_transparency_gradient])
        assert legend.FillTransparenceGradientName.startswith("ChartTransparencyGradient")

        font_size100 = UnitMM100.from_pt(11.3)
        legend_font_only_style = LegendFontOnly(size=font_size100)
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_font_only_style])
        assert UnitMM100.from_pt(legend.CharHeight).value in range(
            font_size100.value - 2, font_size100.value + 3
        )  # +/- 2mm100

        font_size100 = UnitMM100.from_pt(12.2)
        legend_font_style = LegendFont(b=True, color=StandardColor.PURPLE, size=font_size100)
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_font_style])
        assert UnitMM100.from_pt(legend.CharHeight).value in range(
            font_size100.value - 2, font_size100.value + 3
        )  # +/- 2mm100
        assert legend.CharColor == StandardColor.PURPLE

        legend_font_effects_style = LegendFontEffects(color=StandardColor.RED, strike=FontStrikeoutEnum.SINGLE)
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_font_effects_style])
        assert legend.CharStrikeout == FontStrikeoutEnum.SINGLE.value

        legend_pos = ChartLegendPosition(pos=LegendPosition.PAGE_END, no_overlap=True, mode=DirectionModeKind.LR_TB)
        Chart2.style_legend(chart_doc=chart_doc, styles=[legend_pos])
        assert legend.AnchorPosition == LegendPosition.PAGE_END
        assert legend.Overlay == False
        assert legend.WritingMode == DirectionModeKind.LR_TB.value

        pos_x100 = UnitMM100.from_cm(7.12)
        pos_y100 = UnitMM100.from_cm(1.63)
        pos_shape_style = ChartShapePosition(pos_x=pos_x100, pos_y=pos_y100)
        Chart2.style_legend(chart_doc=chart_doc, styles=[pos_shape_style])
        legend_point = cast("UnoPoint", chart_doc.Legend.Position)
        assert legend_point.X in range(pos_x100.value - 2, pos_x100.value + 3)  # +/- 2mm100
        assert legend_point.Y in range(pos_y100.value - 2, pos_y100.value + 3)  # +/- 2mm100

        width_100 = UnitMM100.from_cm(4.1)
        height_100 = UnitMM100.from_cm(3.5)
        size_shape_style = ChartShapeSize(width=width_100, height=height_100)
        Chart2.style_legend(chart_doc=chart_doc, styles=[size_shape_style])

        legend_size = cast("UnoSize", chart_doc.Legend.Size)
        assert legend_size.Width in range(width_100.value - 2, width_100.value + 3)  # +/- 2mm100
        assert legend_size.Height in range(height_100.value - 2, height_100.value + 3)  # +/- 2mm100

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
