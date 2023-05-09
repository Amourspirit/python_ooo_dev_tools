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
from ooodev.format.chart2.direct.title.area import (
    Color as ChartTitleBgColor,
    Gradient as ChartTitleBgGradient,
    Hatch as ChartTitleBgHatch,
    Img as ChartTitleBgImg,
    Pattern as ChartTitleBgPattern,
    PresetGradientKind,
    PresetHatchKind,
    PresetImageKind,
    PresetPatternKind,
)
from ooodev.format.chart2.direct.title.font import Font as TitleFont
from ooodev.format.chart2.direct.title.borders import LineProperties as TitleBorderLineProperties, BorderLineKind

from ooodev.format.chart2.direct.title.position_size import Position as TitlePosition

from ooodev.utils.color import CommonColor, StandardColor
from ooodev.utils.info import Info
from ooodev.units import UnitMM100

if TYPE_CHECKING:
    from com.sun.star.chart2 import Title
    from com.sun.star.style import CharacterProperties


def test_calc_set_styles_subtitle(loader, copy_fix_calc) -> None:
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
        title_font = TitleFont(b=True, size=16, color=CommonColor.DARK_BLUE)
        title_area_bg_color = ChartTitleBgColor(CommonColor.LIGHT_GREEN)
        Chart2.set_title(
            chart_doc=chart_doc,
            title=Calc.get_string(sheet=sheet, cell_name="A1"),
            styles=[title_font, title_area_bg_color],
        )

        subtitle_font = TitleFont(i=True, size=12, color=StandardColor.RED_LIGHT2)
        subtitle_area_bg_color = ChartTitleBgColor(StandardColor.RED_DARK2)
        Chart2.set_subtitle(
            chart_doc=chart_doc, subtitle="Sales by month", styles=[subtitle_font, subtitle_area_bg_color]
        )

        Chart2.set_x_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A2"))
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B2"))

        Calc.goto_cell(cell_name="A1", doc=doc)

        subtitle = cast("Title", Chart2.get_subtitle(chart_doc=chart_doc))
        subtitle.FillColor == StandardColor.RED_DARK2

        fo_strs = subtitle.getText()
        fo_first = cast("CharacterProperties", fo_strs[0])
        assert fo_first.CharColor == StandardColor.RED_LIGHT2
        assert fo_first.CharHeight == 12

        grad = ChartTitleBgGradient.from_preset(chart_doc, PresetGradientKind.NEON_LIGHT)
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[grad])
        subtitle = cast("Title", Chart2.get_subtitle(chart_doc=chart_doc))
        assert subtitle.FillGradientName == PresetGradientKind.NEON_LIGHT.value

        hatch = ChartTitleBgHatch(
            chart_doc=chart_doc, color=CommonColor.INDIGO, space=3, angle=40, bg_color=CommonColor.LIGHT_SKY_BLUE
        )
        subtitle_border = TitleBorderLineProperties(
            style=BorderLineKind.CONTINUOUS, width=0.5, color=CommonColor.BLUE_VIOLET, transparency=15
        )
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[hatch, subtitle_border])
        subtitle = cast("Title", Chart2.get_subtitle(chart_doc=chart_doc))
        assert subtitle.FillHatchName == hatch.prop_hatch_name
        assert subtitle.FillBackground
        assert subtitle.FillColor == CommonColor.LIGHT_SKY_BLUE
        assert subtitle.LineColor == CommonColor.BLUE_VIOLET
        assert subtitle.LineTransparence == 15

        img = ChartTitleBgImg.from_preset(chart_doc, PresetImageKind.COFFEE_BEANS)
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[img])
        subtitle = cast("Title", Chart2.get_subtitle(chart_doc=chart_doc))
        assert subtitle.FillBitmapName == str(PresetImageKind.COFFEE_BEANS)

        subtitle_border = TitleBorderLineProperties(style=BorderLineKind.NONE)
        pattern = ChartTitleBgPattern.from_preset(chart_doc, PresetPatternKind.DIAGONAL_BRICK)
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[pattern, subtitle_border])
        subtitle = cast("Title", Chart2.get_subtitle(chart_doc=chart_doc))
        assert subtitle.FillBitmapName == str(PresetPatternKind.DIAGONAL_BRICK)
        assert subtitle.LineStyle is None

        pos_x = UnitMM100.from_cm(3.94)
        pos_y = UnitMM100.from_cm(1.79)
        subtitle_pos = TitlePosition(pos_x=pos_x, pos_y=pos_y)
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[subtitle_pos])
        subtitle_shape = chart_doc.getSubTitle()
        shape_pos = subtitle_shape.getPosition()
        assert shape_pos.X in range(pos_x.value - 2, pos_x.value + 3)  # +/- 2 mm100 tolerance

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
