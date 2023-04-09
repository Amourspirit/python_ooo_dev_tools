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
from ooodev.format.chart.direct.title.area import (
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
from ooodev.format.chart.direct.title.font import Font as TitleFont
from ooodev.utils.color import CommonColor, StandardColor
from ooodev.utils.info import Info

if TYPE_CHECKING:
    from com.sun.star.chart2 import Title
    from com.sun.star.style import CharacterProperties


def test_calc_set_title_styles(loader, copy_fix_calc) -> None:
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
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A2"))
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B2"))

        Calc.goto_cell(cell_name="A1", doc=doc)

        title = cast("Title", Chart2.get_title(chart_doc=chart_doc))
        title.FillColor == CommonColor.LIGHT_GREEN

        fo_strs = title.getText()
        fo_first = cast("CharacterProperties", fo_strs[0])
        assert fo_first.CharColor == CommonColor.DARK_BLUE
        assert fo_first.CharHeight == 16

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_calc_set_styles_title(loader, copy_fix_calc) -> None:
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
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A2"))
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B2"))

        title_area_bg_color = ChartTitleBgColor(CommonColor.LIGHT_YELLOW)
        title_font = TitleFont(b=True, size=14, color=CommonColor.DARK_GREEN)
        Chart2.style_title(chart_doc=chart_doc, styles=[title_area_bg_color, title_font])

        Calc.goto_cell(cell_name="A1", doc=doc)

        title = cast("Title", Chart2.get_title(chart_doc=chart_doc))
        title.FillColor == CommonColor.LIGHT_YELLOW

        fo_strs = title.getText()
        fo_first = cast("CharacterProperties", fo_strs[0])
        assert fo_first.CharColor == CommonColor.DARK_GREEN
        assert fo_first.CharHeight == 14

        grad = ChartTitleBgGradient.from_preset(PresetGradientKind.NEON_LIGHT)
        Chart2.style_title(chart_doc=chart_doc, styles=[grad])
        title = cast("Title", Chart2.get_title(chart_doc=chart_doc))
        assert title.FillGradientName == PresetGradientKind.NEON_LIGHT.value

        hatch = ChartTitleBgHatch.from_preset(chart_doc, PresetHatchKind.GREEN_90_DEGREES_TRIPLE)
        Chart2.style_title(chart_doc=chart_doc, styles=[hatch])
        title = cast("Title", Chart2.get_title(chart_doc=chart_doc))
        assert title.FillHatchName == PresetHatchKind.GREEN_90_DEGREES_TRIPLE.value
        assert title.FillHatchName == hatch.prop_hatch_name

        hatch = ChartTitleBgHatch(
            chart_doc=chart_doc, color=CommonColor.INDIGO, space=3, angle=40, bg_color=CommonColor.LIGHT_SKY_BLUE
        )
        Chart2.style_title(chart_doc=chart_doc, styles=[hatch])
        title = cast("Title", Chart2.get_title(chart_doc=chart_doc))
        assert title.FillHatchName == hatch.prop_hatch_name
        assert title.FillBackground
        assert title.FillColor == CommonColor.LIGHT_SKY_BLUE

        # chart_doc_ms_factory = Lo.qi(XMultiServiceFactory, chart_doc)
        # container = Lo.create_instance_msf(XNameContainer, "com.sun.star.drawing.HatchTable", msf=chart_doc_ms_factory)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


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

        grad = ChartTitleBgGradient.from_preset(PresetGradientKind.NEON_LIGHT)
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[grad])
        subtitle = cast("Title", Chart2.get_subtitle(chart_doc=chart_doc))
        assert subtitle.FillGradientName == PresetGradientKind.NEON_LIGHT.value

        hatch = ChartTitleBgHatch(
            chart_doc=chart_doc, color=CommonColor.INDIGO, space=3, angle=40, bg_color=CommonColor.LIGHT_SKY_BLUE
        )
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[hatch])
        subtitle = cast("Title", Chart2.get_subtitle(chart_doc=chart_doc))
        assert subtitle.FillHatchName == hatch.prop_hatch_name
        assert subtitle.FillBackground
        assert subtitle.FillColor == CommonColor.LIGHT_SKY_BLUE

        img = ChartTitleBgImg.from_preset(chart_doc, PresetImageKind.COFFEE_BEANS)
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[img])
        subtitle = cast("Title", Chart2.get_subtitle(chart_doc=chart_doc))
        assert subtitle.FillBitmapName == str(PresetImageKind.COFFEE_BEANS)

        pattern = ChartTitleBgPattern.from_preset(chart_doc, PresetPatternKind.DIAGONAL_BRICK)
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[pattern])
        subtitle = cast("Title", Chart2.get_subtitle(chart_doc=chart_doc))
        assert subtitle.FillBitmapName == str(PresetPatternKind.DIAGONAL_BRICK)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
