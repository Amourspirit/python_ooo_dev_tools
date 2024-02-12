from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])
try:
    from ooodev.office.chart2 import Chart2
except ImportError:
    Chart2 = None
from ooodev.utils.info import Info
from ooodev.utils.kind.chart2_types import ChartTypes
from ooodev.utils.color import StandardColor


def test_insert_chart2(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5) or Chart2 is None:
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    fix_path = cast(Path, copy_fix_calc("chartsData.ods"))

    # get_sheet is overload method.
    # testing each overload.
    from ooodev.calc import CalcDoc

    doc = CalcDoc.open_doc(fix_path)
    try:
        sheet = doc.sheets[0]
        rng_obj = sheet.rng("A2:B8")
        chart = sheet.charts.insert_chart(
            rng_obj=rng_obj, diagram_name=ChartTypes.Column.DEFAULT, color_bg=None, color_wall=None
        )
        assert chart is not None
        chart_doc = chart.chart_doc

        assert chart_doc is not None
        draw_page = chart.draw_page
        assert draw_page is not None

        chart_title = chart_doc.get_title()
        assert chart_title is None
        count = 0
        for cht in sheet.charts:
            count += 1
            assert cht is not None
        assert count == 1

        assert len(sheet.charts) == 1

        sheet.charts.clear()
        assert len(sheet.charts) == 0

        rng = sheet.get_range(range_obj=rng_obj)
        chart = rng.insert_chart(diagram_name=ChartTypes.Column.DEFAULT, color_bg=None, color_wall=None)
        assert chart is not None
        chart_doc = chart.chart_doc

        assert chart_doc is not None
        draw_page = chart.draw_page
        assert draw_page is not None

        chart_title = chart_doc.get_title()
        assert chart_title is None
        count = 0
        for cht in sheet.charts:
            count += 1
            assert cht is not None
        assert count == 1

        assert len(sheet.charts) == 1

        sheet.charts.clear()
        assert len(sheet.charts) == 0

    finally:
        doc.close()


def test_chart2(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5) or Chart2 is None:
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
    from ooo.dyn.chart2.legend_position import LegendPosition
    from ooo.dyn.chart.chart_axis_position import ChartAxisPosition
    from ooodev.format.inner.preset.preset_border_line import BorderLineKind
    from ooodev.format.inner.direct.write.char.font.font_effects import FontLine, FontUnderlineEnum
    from ooodev.format.inner.direct.chart2.title.alignment.direction import DirectionModeKind

    # from ooodev.format.inner.preset.preset_gradient import PresetGradientKind
    # from ooodev.format.inner.preset.preset_hatch import PresetHatchKind
    # from ooodev.format.inner.preset.preset_image import PresetImageKind
    from ooodev.format.inner.preset.preset_pattern import PresetPatternKind
    from ooodev.units import Angle
    from ooodev.utils.data_type.intensity_range import IntensityRange

    fix_path = cast(Path, copy_fix_calc("chartsData.ods"))

    # get_sheet is overload method.
    # testing each overload.
    from ooodev.calc import CalcDoc

    doc = CalcDoc.open_doc(fix_path)
    try:
        sheet = doc.sheets[0]
        rng_obj = sheet.rng("A2:B8")
        chart = sheet.charts.insert_chart(
            rng_obj=rng_obj, diagram_name=ChartTypes.Column.DEFAULT, color_bg=None, color_wall=None
        )
        assert chart is not None
        chart_doc = chart.chart_doc

        assert chart_doc is not None

        # lock and unlock controllers
        with chart_doc:
            chart_title = chart_doc.set_title("Chart Title")
            assert chart_title is not None
            style = chart_title.style_font_effect(color=StandardColor.BRICK_DARK2)
            assert style is not None

        chart_title = chart_doc.get_title()
        assert chart_title is not None
        chart_title.rotation = 10
        assert chart_title.rotation.value == 10

        style = chart_title.style_position(3, 3)
        assert style is not None

        assert chart_doc.axis_x is not None
        assert chart_doc.axis_y is not None

        style = chart_doc.axis_y.style_numbers_numbers(
            source_format=False, num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2
        )
        assert style is not None

        x_title = chart_doc.axis_x.get_title()
        assert x_title is None

        y_title = chart_doc.axis_y.get_title()
        assert y_title is None

        x_title = chart_doc.axis_x.set_title("X Axis Title")
        assert x_title is not None

        style = x_title.style_area_color(StandardColor.BLUE_LIGHT2)
        assert style is not None

        # style = x_title.style_area_gradient_from_preset(preset=PresetGradientKind.MAHOGANY)
        # assert style is not None

        # style = x_title.style_area_hatch_from_preset(preset=PresetHatchKind.BLUE_45_DEGREES_CROSSED)
        # assert style is not None

        # style = x_title.style_area_image_from_preset(preset=PresetImageKind.BRICK_WALL)
        # assert style is not None

        # style = x_title.style_area_pattern_from_preset(preset=PresetPatternKind.SHINGLE)
        # assert style is not None

        style = x_title.style_border_line(color=StandardColor.GREEN_DARK2, width=0.8)
        assert style is not None

        y_title = chart_doc.axis_y.set_title("Y Axis Title")
        assert y_title is not None

        x_title = chart_doc.axis_x.get_title()
        assert x_title is not None
        y_title = chart_doc.axis_y.get_title()
        assert y_title is not None

        y_title.rotation = 15
        assert y_title.rotation.value == 15

        style = y_title.style_orientation(angle=0, vertical=True)
        assert style is not None

        chart_doc.axis_y.style_gird_line(style=BorderLineKind.CONTINUOUS, color=StandardColor.RED, width=0.5)

        # chart_doc.set_bg_color(StandardColor.GRAY_LIGHT2)
        chart_doc.style_area_color(StandardColor.BRICK_LIGHT2)
        # chart_doc.style_area_gradient_from_preset(preset=PresetGradientKind.MAHOGANY)
        # chart_doc.style_area_image_from_preset(preset=PresetImageKind.BRICK_WALL)
        # chart_doc.style_area_pattern_from_preset(preset=PresetPatternKind.SHINGLE)
        # style = chart_doc.style_area_hatch_from_preset(preset=PresetHatchKind.BLUE_45_DEGREES_CROSSED)
        # style = chart_doc.style_area_transparency_transparency(50)
        style = chart_doc.style_area_transparency_gradient(angle=Angle(30), grad_intensity=IntensityRange(0, 100))
        assert style is not None

        chart_doc.set_wall_color(StandardColor.RED_LIGHT2)

        chart_doc.axis_x.style_font_effect(color=StandardColor.RED_DARK2)
        chart_doc.axis_y.style_font_effect(color=StandardColor.GREEN_DARK2)

        x_title.style_font_effect(color=StandardColor.BLUE_DARK2)
        # y_title.style_font_effect(color=StandardColor.LIME_DARK3)
        y_font_effect = y_title.style_font_effect()
        if y_font_effect is not None:
            y_font_effect.fmt_color(StandardColor.LIME_DARK3).fmt_shadowed(True).update()

        coord_sys = chart_doc.first_diagram.get_coordinate_system()
        assert coord_sys is not None

        chart_type = coord_sys.chart_type
        assert chart_type is not None
        data_series = chart_type.get_data_series()
        assert len(data_series) > 0
        series = data_series[0]
        assert series is not None

        # style = series.style_numbers_numbers(
        #     source_format=False, num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2
        # )
        # assert style is not None

        data_series[0].style_font_effect(color=StandardColor.MAGENTA_DARK2)
        style = data_series[0].style_font(name="Always In My Heart", font_style="Bold Italic")
        assert style is not None
        style = style.fmt_size(16)
        style.update()

        style = data_series[0].style_area_color(StandardColor.GREEN_LIGHT2)
        assert style is not None

        # style = data_series[0].style_area_gradient_from_preset(preset=PresetGradientKind.MAHOGANY)
        # assert style is not None

        # style = data_series[0].style_area_hatch_from_preset(preset=PresetHatchKind.BLUE_45_DEGREES_CROSSED)
        # assert style is not None

        # style = data_series[0].style_area_image_from_preset(preset=PresetImageKind.BRICK_WALL)
        # assert style is not None

        # style = data_series[0].style_area_pattern_from_preset(preset=PresetPatternKind.SHINGLE)
        # assert style is not None

        style = data_series[0].style_border_line(color=StandardColor.BLUE_DARK2, width=0.8)
        assert style is not None

        # style = data_series[0].style_area_transparency_transparency(50)
        # assert style is not None

        # style = data_series[0].style_area_transparency_gradient(angle=Angle(30), grad_intensity=IntensityRange(0, 100))
        # assert style is not None

        style = data_series[0].style_label_border_line(color=StandardColor.BLUE_DARK2, width=0.8)
        assert style is not None

        dp = data_series[0][2]
        style = dp.style_font_effect(color=StandardColor.GREEN_DARK3)
        assert style is not None

        style = dp.style_area_color(StandardColor.BLUE_LIGHT2)
        assert style is not None

        # style = dp.style_area_gradient_from_preset(preset=PresetGradientKind.NEON_LIGHT)
        # assert style is not None

        # style = dp.style_area_hatch_from_preset(preset=PresetHatchKind.GREEN_90_DEGREES_TRIPLE)
        # assert style is not None

        # style = dp.style_area_image_from_preset(preset=PresetImageKind.ICE_LIGHT)
        # assert style is not None

        # style = dp.style_area_pattern_from_preset(preset=PresetPatternKind.ZIG_ZAG)
        # assert style is not None

        style = dp.style_border_line(color=StandardColor.RED_DARK2, width=1.1)
        assert style is not None

        # style = dp.style_area_transparency_transparency(20)
        # assert style is not None

        # style = dp.style_area_transparency_gradient(angle=Angle(45), grad_intensity=IntensityRange(100, 0))
        # assert style is not None

        style = dp.style_label_border_line(color=StandardColor.INDIGO_DARK3, width=0.5)
        assert style is not None

        style = dp.style_numbers_numbers(source_format=False, num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2)
        assert style is not None

        ds = data_series[0].get_data_source()
        assert ds is not None

        chart_doc.axis_x.style_axis_line(color=StandardColor.BLUE_DARK2)
        style = chart_doc.axis_x.style_axis_pos_axis_line(cross=ChartAxisPosition.VALUE, value=40)
        assert style is not None

        x_title = chart_doc.axis_x.get_title()
        assert x_title is not None
        style = x_title.style_position(3, 3)
        # can't set position for axis title
        assert style is None

        axis2_x = chart_doc.axis2_x
        if axis2_x is not None:
            axis2_x.set_title("X2 Axis Title")

        chart_img = chart.shape.get_image()
        assert chart_img is not None

        sub_title = chart_doc.first_diagram.get_title()
        assert sub_title is None

        style = chart_doc.first_diagram.wall.style_area_color(StandardColor.YELLOW_LIGHT2)
        assert style is not None

        # style = chart_doc.first_diagram.wall.style_area_gradient_from_preset(preset=PresetGradientKind.MAHOGANY)
        # assert style is not None

        # style = chart_doc.first_diagram.wall.style_area_hatch_from_preset(
        #     preset=PresetHatchKind.BLUE_45_DEGREES_CROSSED
        # )
        # assert style is not None

        # style = chart_doc.first_diagram.wall.style_area_image_from_preset(preset=PresetImageKind.BRICK_WALL)
        # assert style is not None

        # style = chart_doc.first_diagram.wall.style_area_pattern_from_preset(preset=PresetPatternKind.SHINGLE)
        # assert style is not None

        style = chart_doc.first_diagram.wall.style_border_line(color=StandardColor.GREEN_DARK3, width=1.5)
        assert style is not None

        # style = chart_doc.first_diagram.wall.style_area_transparency_transparency(66)
        # assert style is not None

        style = chart_doc.first_diagram.wall.style_area_transparency_gradient(
            angle=Angle(40), grad_intensity=IntensityRange(0, 100)
        )
        assert style is not None

        sub_title = chart_doc.first_diagram.set_title("Sub Title")
        assert sub_title is not None

        style = sub_title.style_position(3, 20)
        assert style is not None

        sub_title = chart_doc.first_diagram.get_title()
        assert sub_title is not None
        sub_title.style_font_effect(color=StandardColor.MAGENTA_LIGHT1)

        style = sub_title.style_position(10, 30)
        assert style is not None

        # pprint(chart_doc.get_templates())

        legend = chart_doc.first_diagram.get_legend()
        assert legend is None

        chart_doc.first_diagram.view_legend(True)
        legend = chart_doc.first_diagram.get_legend()
        assert legend is not None

        legend.style_font_effect(
            color=StandardColor.PURPLE,
            underline=FontLine(line=FontUnderlineEnum.BOLDWAVE, color=StandardColor.GREEN_DARK2),
        )
        legend.style_font(size=16)

        style = legend.style_area_transparency_transparency(0)
        assert style is not None

        # needs area transparency to be turned off.
        style = legend.style_area_color(StandardColor.YELLOW_LIGHT2)
        assert style is not None

        # style = legend.style_area_transparency_gradient(angle=Angle(40), grad_intensity=IntensityRange(0, 100))
        # assert style is not None

        # style = legend.style_area_gradient_from_preset(preset=PresetGradientKind.MAHOGANY)
        # assert style is not None

        # style = legend.style_area_hatch_from_preset(preset=PresetHatchKind.BLUE_45_DEGREES_CROSSED)
        # assert style is not None

        # style = legend.style_area_image_from_preset(preset=PresetImageKind.BRICK_WALL)
        # assert style is not None

        style = legend.style_area_pattern_from_preset(preset=PresetPatternKind.SHINGLE)
        assert style is not None

        style = legend.style_border_line(color=StandardColor.BLUE_DARK2, width=1.5)
        assert style is not None

        style = legend.style_position(pos=LegendPosition.PAGE_END, no_overlap=True, mode=DirectionModeKind.LR_TB)
        assert style is not None

        bdr_style = chart_doc.style_border_line(color=StandardColor.GREEN_DARK3)
        assert bdr_style is not None
        bdr_style.prop_width = 2.2
        assert bdr_style.update()

        style = chart.shape.style_position(10, 10)
        assert style is not None

        style = chart.shape.style_size(300, 150)
        assert style is not None

        # chart_doc.first_diagram.view_legend(False)

        # chart.shape.export_shape_png(Path.cwd() / "tmp" / "chart.png", 200)

        # del sheet.charts[0]
        sheet.charts.clear()
        assert len(sheet.charts) == 0

    finally:
        doc.close()
