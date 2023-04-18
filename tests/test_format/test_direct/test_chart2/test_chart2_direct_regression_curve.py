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
    from ooodev.office.chart2 import Chart2, CurveKind
except ImportError:
    Chart2 = None
    CurveKind = None

from ooodev.utils.data_type.angle import Angle
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
from ooodev.format.chart2.direct.numbers import Numbers, NumberFormatIndexEnum

from ooodev.utils.color import CommonColor, StandardColor
from ooodev.utils.info import Info
from ooodev.units import UnitMM100
from ooodev.format.chart2.direct.title.alignment import Orientation as TitleOrientation
from ooodev.format.chart2.direct.title.alignment import Direction as TitleDirection, DirectionModeKind

if TYPE_CHECKING:
    from com.sun.star.chart2 import Title
    from com.sun.star.style import CharacterProperties


def test_calc_set_styles_title(loader, copy_fix_calc) -> None:
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

        range_addr = Calc.get_address(sheet=sheet, range_name="A110:B122")
        chart_doc = Chart2.insert_chart(
            sheet=sheet,
            cells_range=range_addr,
            cell_name="C109",
            width=16,
            height=11,
            diagram_name=ChartTypes.XY.TEMPLATE_LINE.SCATTER_SYMBOL,
        )
        Calc.goto_cell(cell_name="A104", doc=doc)

        Chart2.set_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A109"))
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A110"))
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B110"))
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))

        title_area_bg_color = ChartTitleBgColor(CommonColor.LIGHT_YELLOW)
        title_font = TitleFont(b=True, size=14, color=CommonColor.DARK_GREEN)
        title_border = TitleBorderLineProperties(style=BorderLineKind.DASH_DOT, width=1.0, color=CommonColor.DARK_RED)
        num_style = Numbers(chart_doc, num_format_index=NumberFormatIndexEnum.NUMBER_DEC2)
        _ = Chart2.draw_regression_curve(
            chart_doc=chart_doc,
            curve_kind=CurveKind.LINEAR,
            styles=[title_area_bg_color, title_font, title_border, num_style],
        )

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
