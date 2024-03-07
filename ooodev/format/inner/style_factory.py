from __future__ import annotations
from typing import Any, TYPE_CHECKING, Type

# pylint: disable=redefined-outer-name
from ooodev.mock import mock_g

if TYPE_CHECKING:
    from ooodev.format.proto.font.font_effects_t import FontEffectsT
    from ooodev.format.proto.font.font_only_t import FontOnlyT
    from ooodev.format.proto.area.fill_color_t import FillColorT
    from ooodev.format.proto.write.char.font.font_t import FontT
    from ooodev.format.proto.chart2.area.chart_fill_gradient_t import ChartFillGradientT

    # from ooodev.format.proto.area.fill_img_t import FillImgT as AreaFillImgT
    from ooodev.format.proto.write.fill.area.fill_img_t import FillImgT as WriteFillImgT
    from ooodev.format.proto.chart2.area.chart_fill_img_t import ChartFillImgT
    from ooodev.format.proto.chart2.area.chart_fill_pattern_t import ChartFillPatternT
    from ooodev.format.proto.chart2.area.chart_fill_hatch_t import ChartFillHatchT
    from ooodev.format.proto.chart2.position_size.position_t import PositionT as Chart2PositionT
    from ooodev.format.proto.chart2.position_size.size_t import SizeT as Chart2SizeT
    from ooodev.format.proto.draw.position_size.size_t import SizeT as DrawSizeT
    from ooodev.format.proto.draw.position_size.position_t import PositionT as DrawPositionT
    from ooodev.format.proto.draw.position_size.protect_t import ProtectT as DrawProtectT
    from ooodev.format.proto.borders.line_properties_t import LinePropertiesT as BorderLinePropertiesT
    from ooodev.format.proto.area.transparency.transparency_t import TransparencyT as TransparencyTransparencyT
    from ooodev.format.proto.area.transparency.gradient_t import GradientT as TransparencyGradientT
    from ooodev.format.proto.chart2.axis.positioning.axis_line_t import AxisLineT as Chart2AxisLineT
    from ooodev.format.proto.chart2.axis.positioning.interval_marks_t import IntervalMarksT as Chart2IntervalMarksT
    from ooodev.format.proto.chart2.axis.positioning.label_position_t import LabelPositionT as Chart2LabelPositionT
    from ooodev.format.proto.chart2.axis.positioning.position_axis_t import PositionAxisT as Chart2PositionAxisT
    from ooodev.format.proto.calc.numbers.numbers_t import NumbersT as CalcNumbersT
    from ooodev.format.proto.chart2.numbers.numbers_t import NumbersT as Chart2NumbersT
    from ooodev.format.proto.chart2.series.data_labels.data_labels.numbers_t import (
        NumbersT as Chart2SeriesDataLabelsNumbersT,
    )
    from ooodev.format.proto.calc.alignment.text_align_t import TextAlignT as CalcAlignTextT
    from ooodev.format.proto.calc.alignment.text_orientation_t import TextOrientationT as CalcAlignOrientationT
    from ooodev.format.proto.calc.alignment.properties_t import PropertiesT as CalcAlignPropertiesT
    from ooodev.format.proto.calc.borders.borders_t import BordersT as CalcBordersT
    from ooodev.format.proto.write.char.font.font_position_t import FontPositionT as WriteCharFontPositionT
    from ooodev.format.proto.font.highlight_t import HighlightT as WriteCharFontHighlightT
else:
    FontEffectsT = Any
    FontOnlyT = Any
    FillColorT = Any
    ChartFillGradientT = Any
    ChartFillImgT = Any
    ChartFillHatchT = Any
    Chart2PositionT = Any
    Chart2SizeT = Any
    DrawSizeT = Any
    DrawPositionT = Any
    DrawProtectT = Any
    BorderLinePropertiesT = Any
    TransparencyTransparencyT = Any
    TransparencyGradientT = Any
    Chart2AxisLineT = Any
    Chart2IntervalMarksT = Any
    Chart2LabelPositionT = Any
    Chart2PositionAxisT = Any
    CalcNumbersT = Any
    Chart2NumbersT = Any
    Chart2SeriesDataLabelsNumbersT = Any
    CalcAlignTextT = Any
    CalcAlignOrientationT = Any
    CalcAlignPropertiesT = Any
    CalcBordersT = Any
    WriteCharFontPositionT = Any
    WriteFillImgT = Any

# pylint: disable=import-outside-toplevel


# region Font
def font_only_factory(name: str) -> Type[FontOnlyT]:
    if name == "ooodev.write.char":
        from ooodev.format.inner.direct.write.char.font.font_only import FontOnly as FontOnly1

        return FontOnly1

    if name == "ooodev.chart2.axis":
        from ooodev.format.inner.direct.chart2.axis.font.font_only import FontOnly as FontOnly2

        return FontOnly2

    if name == "ooodev.chart2.title":
        from ooodev.format.inner.direct.chart2.title.font.font_only import FontOnly as FontOnly3

        return FontOnly3

    if name == "ooodev.chart2.series.data_labels":
        from ooodev.format.inner.direct.chart2.series.data_labels.font.font_only import FontOnly as FontOnly4

        return FontOnly4

    if name == "ooodev.chart2.legend":
        from ooodev.format.inner.direct.chart2.legend.font.font_only import FontOnly as FontOnly5

        return FontOnly5

    if name in {"ooodev.calc.cell", "ooodev.calc.cell_rng"}:
        from ooodev.format.inner.direct.calc.char.font.font_only import FontOnly as FontOnly6

        return FontOnly6

    raise ValueError(f"Invalid name: {name}")


def font_factory(name: str) -> Type[FontT]:
    if name == "ooodev.write.char":
        from ooodev.format.inner.direct.write.char.font.font import Font as Font1

        return Font1

    if name == "ooodev.general_style.text":
        from ooodev.format.inner.direct.general_style.text.font import Font as Font2

        return Font2

    raise ValueError(f"Invalid name: {name}")


def font_effects_factory(name: str) -> Type[FontEffectsT]:
    if name == "ooodev.write.char":
        from ooodev.format.inner.direct.write.char.font.font_effects import FontEffects as FontEffects1

        return FontEffects1
    if name == "ooodev.chart2.axis":
        from ooodev.format.inner.direct.chart2.axis.font.font_effects import FontEffects as FontEffects2

        return FontEffects2

    if name == "ooodev.chart2.title":
        from ooodev.format.inner.direct.chart2.title.font.font_effects import FontEffects as FontEffects3

        return FontEffects3

    if name == "ooodev.chart2.series.data_labels":
        from ooodev.format.inner.direct.chart2.series.data_labels.font.font_effects import FontEffects as FontEffects4

        return FontEffects4

    if name == "ooodev.chart2.legend":
        from ooodev.format.inner.direct.chart2.legend.font.font_effects import FontEffects as FontEffects5

        return FontEffects5

    if name in {"ooodev.calc.cell", "ooodev.calc.cell_rng"}:
        from ooodev.format.inner.direct.calc.char.font.font_effects import FontEffects as FontEffects6

        return FontEffects6

    raise ValueError(f"Invalid name: {name}")


def font_position_factory(name: str) -> Type[WriteCharFontPositionT]:
    if name == "ooodev.write.char":
        from ooodev.format.inner.direct.write.char.font.font_position import FontPosition as FontPosition1

        return FontPosition1

    raise ValueError(f"Invalid name: {name}")


# endregion Font


# region Calc
def calc_align_text_factory(name: str) -> Type[CalcAlignTextT]:
    if name in {"ooodev.calc.cell", "ooodev.calc.cell_rng"}:
        from ooodev.format.inner.direct.calc.alignment.text_align import TextAlign as TextAlign1

        return TextAlign1

    raise ValueError(f"Invalid name: {name}")


def calc_align_orientation_factory(name: str) -> Type[CalcAlignOrientationT]:
    if name in {"ooodev.calc.cell", "ooodev.calc.cell_rng"}:
        from ooodev.format.inner.direct.calc.alignment.text_orientation import TextOrientation as TextOrientation1

        return TextOrientation1

    raise ValueError(f"Invalid name: {name}")


def calc_align_properties_factory(name: str) -> Type[CalcAlignPropertiesT]:
    if name in {"ooodev.calc.cell", "ooodev.calc.cell_rng"}:
        from ooodev.format.inner.direct.calc.alignment.properties import Properties as Properties1

        return Properties1

    raise ValueError(f"Invalid name: {name}")


def calc_borders_factory(name: str) -> Type[CalcBordersT]:
    if name in {"ooodev.calc.cell", "ooodev.calc.cell_rng"}:
        from ooodev.format.inner.direct.calc.border.borders import Borders as Borders1

        return Borders1

    raise ValueError(f"Invalid name: {name}")


# endregion Calc


# region numbers format
# region general numbers format
def numbers_numbers_factory(name: str) -> Type[CalcNumbersT]:
    if name == "ooodev.number.numbers":
        from ooodev.format.inner.direct.calc.numbers.numbers import Numbers as Numbers1

        return Numbers1

    raise ValueError(f"Invalid name: {name}")


# endregion general numbers format


# region chart2 numbers format
def chart2_numbers_numbers_factory(name: str) -> Type[Chart2NumbersT]:
    if name == "ooodev.chart2.number.numbers":
        from ooodev.format.inner.direct.chart2.chart.numbers.numbers import Numbers as Numbers2

        return Numbers2

    raise ValueError(f"Invalid name: {name}")


def chart2_series_data_labels_numbers_factory(name: str) -> Type[Chart2SeriesDataLabelsNumbersT]:
    if name == "ooodev.chart2.series.data_labels.number.numbers":
        from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.number_format import (
            NumberFormat as NumberFormat1,
        )

        return NumberFormat1

    if name == "ooodev.chart2.axis.numbers.numbers":
        from ooodev.format.inner.direct.chart2.axis.numbers.numbers import Numbers as NumberFormat2

        return NumberFormat2

    raise ValueError(f"Invalid name: {name}")


# endregion chart2 numbers format
# endregion numbers format


def area_color_factory(name: str) -> Type[FillColorT]:
    if name == "ooodev.chart2.general":
        from ooodev.format.inner.direct.chart2.chart.area.color import Color as Color1

        return Color1

    if name == "ooodev.char2.legend.area.color":
        from ooodev.format.inner.direct.chart2.legend.area.color import Color as Color2

        return Color2

    if name == "ooodev.char2.wall.area":
        from ooodev.format.inner.direct.chart2.wall.area.color import Color as Color3

        return Color3

    if name in {"ooodev.char2.series.data_series.area", "ooodev.char2.series.data_point.area"}:
        from ooodev.format.inner.direct.chart2.series.data_series.area.color import Color as Color4

        return Color4

    if name == "ooodev.char2.title.area":
        from ooodev.format.inner.direct.chart2.title.area.color import Color as Color5

        return Color5

    if name in {"ooodev.calc.cell", "ooodev.calc.cell_rng"}:
        from ooodev.format.inner.direct.calc.background.color import Color as Color6

        return Color6

    if name == "ooodev.write.table.background":
        from ooodev.format.inner.direct.write.table.background.color import Color as Color7

        return Color7
    raise ValueError(f"Invalid name: {name}")


def area_transparency_transparency_factory(name: str) -> Type[TransparencyTransparencyT]:
    if name == "ooodev.area.transparency":
        from ooodev.format.draw.direct.transparency.transparency import Transparency as Transparency1

        return Transparency1
    if name == "ooodev.write.para.transparency":
        from ooodev.format.inner.direct.write.fill.transparent.transparency import Transparency as Transparency2

        return Transparency2

    if name == "ooodev.write.shape.area.transparency":
        from ooodev.format.writer.direct.shape.transparency.transparency import Transparency as Transparency3

        return Transparency3

    if name == "ooodev.chart2.general":
        from ooodev.format.inner.direct.chart2.chart.transparent.transparency import Transparency as Transparency4

        return Transparency4

    if name == "ooodev.chart2.legend":
        from ooodev.format.inner.direct.chart2.legend.transparent.transparency import Transparency as Transparency5

        return Transparency5

    if name == "ooodev.chart2.wall":
        from ooodev.format.inner.direct.chart2.wall.transparent.transparency import Transparency as Transparency6

        return Transparency6

    if name in {"ooodev.char2.series.data_series.transparency", "ooodev.char2.series.data_point.transparency"}:
        from ooodev.format.inner.direct.chart2.series.data_series.transparent.transparency import (
            Transparency as Transparency7,
        )

        return Transparency7
    raise ValueError(f"Invalid name: {name}")


def area_transparency_gradient_factory(name: str) -> Type[TransparencyGradientT]:
    if name == "ooodev.area.transparency":
        from ooodev.format.draw.direct.transparency.gradient import Gradient as Gradient1

        return Gradient1

    if name == "ooodev.write.shape.area.transparency":
        from ooodev.format.writer.direct.shape.transparency.gradient import Gradient as Gradient2

        return Gradient2

    if name == "ooodev.chart2.general":
        from ooodev.format.inner.direct.chart2.chart.transparent.gradient import Gradient as Gradient3

        return Gradient3

    if name == "ooodev.chart2.legend":
        from ooodev.format.inner.direct.chart2.legend.transparent.gradient import Gradient as Gradient4

        return Gradient4

    if name == "ooodev.chart2.wall.transparency":
        from ooodev.format.inner.direct.chart2.wall.transparent.gradient import Gradient as Gradient5

        return Gradient5

    if name in {"ooodev.char2.series.data_series.transparency", "ooodev.char2.series.data_point.transparency"}:
        from ooodev.format.inner.direct.chart2.series.data_series.transparent.gradient import Gradient as Gradient6

        return Gradient6

    raise ValueError(f"Invalid name: {name}")


def write_area_img_factory(name: str) -> Type[WriteFillImgT]:
    if name == "ooodev.write.table.background":
        from ooodev.format.inner.direct.write.table.background.img import Img as Img1

        return Img1

    # if name == "ooodev.chart2.general":
    # ooodev.format.inner.direct.write.table.background.img

    raise ValueError(f"Invalid name: {name}")


# def area_gradient_factory(name: str) -> Type[FillGradientT]:
#     if name == "ooodev.chart2.general":
#         # from ooodev.format.inner.direct.chart2.chart.area.color import Gradient
#         from ooodev.format.inner.direct.chart2.chart.area.gradient import Gradient

#         return Gradient

#     raise ValueError(f"Invalid name: {name}")


def chart2_area_gradient_factory(name: str) -> Type[ChartFillGradientT]:
    if name == "ooodev.chart2.general":
        from ooodev.format.inner.direct.chart2.chart.area.gradient import Gradient as Gradient20

        return Gradient20

    if name == "ooodev.chart2.legend":
        from ooodev.format.inner.direct.chart2.legend.area.gradient import Gradient as Gradient21

        return Gradient21

    if name == "ooodev.char2.wall.area":
        from ooodev.format.inner.direct.chart2.wall.area.gradient import Gradient as Gradient22

        return Gradient22

    if name in {"ooodev.char2.series.data_series.area", "ooodev.char2.series.data_point.area"}:
        from ooodev.format.inner.direct.chart2.series.data_series.area.gradient import Gradient as Gradient23

        return Gradient23

    if name == "ooodev.char2.title":
        from ooodev.format.inner.direct.chart2.title.area.gradient import Gradient as Gradient24

        return Gradient24

    raise ValueError(f"Invalid name: {name}")


def chart2_area_img_factory(name: str) -> Type[ChartFillImgT]:
    if name == "ooodev.chart2.general":
        from ooodev.format.inner.direct.chart2.chart.area.img import Img as Img20

        return Img20

    if name == "ooodev.chart2.legend":
        from ooodev.format.inner.direct.chart2.legend.area.img import Img as Img21

        return Img21

    if name == "ooodev.char2.wall.area":
        from ooodev.format.inner.direct.chart2.wall.area.img import Img as Img22

        return Img22

    if name in {"ooodev.char2.series.data_series.area", "ooodev.char2.series.data_point.area"}:
        from ooodev.format.inner.direct.chart2.series.data_series.area.img import Img as Img23

        return Img23

    if name == "ooodev.char2.title":
        from ooodev.format.inner.direct.chart2.title.area.img import Img as Img24

        return Img24

    # if name == "ooodev.chart2.general":
    # ooodev.format.inner.direct.write.table.background.img

    raise ValueError(f"Invalid name: {name}")


def chart2_area_pattern_factory(name: str) -> Type[ChartFillPatternT]:
    if name == "ooodev.chart2.general":
        from ooodev.format.inner.direct.chart2.chart.area.pattern import Pattern as Pattern20

        return Pattern20

    if name == "ooodev.chart2.legend":
        from ooodev.format.inner.direct.chart2.legend.area.pattern import Pattern as Pattern21

        return Pattern21

    if name == "ooodev.char2.wall.area":
        from ooodev.format.inner.direct.chart2.wall.area.pattern import Pattern as Pattern22

        return Pattern22

    if name in {"ooodev.char2.series.data_series.area", "ooodev.char2.series.data_point.area"}:
        from ooodev.format.inner.direct.chart2.series.data_series.area.pattern import Pattern as Pattern23

        return Pattern23

    if name == "ooodev.char2.title":
        from ooodev.format.inner.direct.chart2.title.area.pattern import Pattern as Pattern24

        return Pattern24

    raise ValueError(f"Invalid name: {name}")


def chart2_area_hatch_factory(name: str) -> Type[ChartFillHatchT]:
    if name == "ooodev.chart2.general":
        from ooodev.format.inner.direct.chart2.chart.area.hatch import Hatch as Hatch20

        return Hatch20

    if name == "ooodev.chart2.legend":
        from ooodev.format.inner.direct.chart2.legend.area.hatch import Hatch as Hatch21

        return Hatch21

    if name == "ooodev.char2.wall.area":
        from ooodev.format.inner.direct.chart2.wall.area.hatch import Hatch as Hatch22

        return Hatch22

    if name in {"ooodev.char2.series.data_series.area", "ooodev.char2.series.data_point.area"}:
        from ooodev.format.inner.direct.chart2.series.data_series.area.hatch import Hatch as Hatch23

        return Hatch23

    if name == "ooodev.char2.title":
        from ooodev.format.inner.direct.chart2.title.area.hatch import Hatch as Hatch24

        return Hatch24

    raise ValueError(f"Invalid name: {name}")


def chart2_position_size_position_factory(name: str) -> Type[Chart2PositionT]:
    if name == "ooodev.chart2.position":
        from ooodev.format.inner.direct.chart2.position_size.position import Position as Position20

        return Position20

    if name == "ooodev.chart2.title":
        from ooodev.format.inner.direct.chart2.title.position_size.position import Position as Position21

        return Position21

    raise ValueError(f"Invalid name: {name}")


def chart2_position_size_size_factory(name: str) -> Type[Chart2SizeT]:
    if name == "ooodev.chart2.size":
        from ooodev.format.inner.direct.chart2.position_size.size import Size as Size20

        return Size20

    raise ValueError(f"Invalid name: {name}")


# region Chart2 Axis
# region Chart2 Axis Positioning
def chart2_axis_pos_line_factory(name: str) -> Type[Chart2AxisLineT]:
    if name == "ooodev.chart2.axis.pos.line":
        from ooodev.format.inner.direct.chart2.axis.positioning.axis_line import AxisLine as AxisLine20

        return AxisLine20

    raise ValueError(f"Invalid name: {name}")


def chart2_axis_pos_interval_factory(name: str) -> Type[Chart2IntervalMarksT]:
    if name == "ooodev.chart2.axis.pos.interval_marks":
        from ooodev.format.inner.direct.chart2.axis.positioning.interval_marks import IntervalMarks as IntervalMarks20

        return IntervalMarks20

    raise ValueError(f"Invalid name: {name}")


def chart2_axis_pos_label_position_factory(name: str) -> Type[Chart2LabelPositionT]:
    if name == "ooodev.chart2.axis.pos.label_position":
        from ooodev.format.inner.direct.chart2.axis.positioning.label_position import LabelPosition as LabelPosition20

        return LabelPosition20

    raise ValueError(f"Invalid name: {name}")


def chart2_axis_pos_position_axis_factory(name: str) -> Type[Chart2PositionAxisT]:
    if name == "ooodev.chart2.axis.pos.position":
        from ooodev.format.inner.direct.chart2.axis.positioning.position_axis import PositionAxis as PositionAxis20

        return PositionAxis20

    raise ValueError(f"Invalid name: {name}")


# endregion Chart2 Axis Positioning

# endregion Chart2 Axis


def draw_position_size_position_factory(name: str) -> Type[DrawPositionT]:
    if name == "ooodev.draw.position":
        from ooodev.format.draw.direct.position_size.position_size.position import Position as Position30

        return Position30

    raise ValueError(f"Invalid name: {name}")


def draw_position_size_size_factory(name: str) -> Type[DrawSizeT]:
    if name == "ooodev.draw.size":
        from ooodev.format.draw.direct.position_size.position_size.size import Size as Size30

        return Size30

    raise ValueError(f"Invalid name: {name}")


def draw_position_size_protect_factory(name: str) -> Type[DrawProtectT]:
    if name == "ooodev.draw.protect":
        from ooodev.format.draw.direct.position_size.position_size.protect import Protect as Protect30

        return Protect30

    raise ValueError(f"Invalid name: {name}")


def draw_border_line_factory(name: str) -> Type[BorderLinePropertiesT]:
    if name == "ooodev.draw.line":
        from ooodev.format.inner.direct.chart2.chart.borders.line_properties import LineProperties as LineProperties30

        return LineProperties30

    if name == "ooodev.chart2.line":
        from ooodev.format.inner.direct.chart2.chart.borders.line_properties import LineProperties as LineProperties31

        return LineProperties31

    if name == "ooodev.chart2.axis.line":
        from ooodev.format.inner.direct.chart2.axis.line.line_properties import LineProperties as LineProperties32

        return LineProperties32

    if name == "ooodev.chart2.grid.line":
        from ooodev.format.inner.direct.chart2.grid.line_properties import LineProperties as LineProperties33

        return LineProperties33

    if name == "ooodev.chart2.legend":
        from ooodev.format.inner.direct.chart2.legend.borders.line_properties import LineProperties as LineProperties34

        return LineProperties34

    if name == "ooodev.chart2.wall.borders":
        from ooodev.format.inner.direct.chart2.wall.borders.line_properties import LineProperties as LineProperties35

        return LineProperties35

    if name in {"ooodev.char2.series.data_series.borders", "ooodev.char2.series.data_point.borders"}:
        from ooodev.format.inner.direct.chart2.series.data_series.borders.line_properties import (
            LineProperties as LineProperties36,
        )

        return LineProperties36

    if name in {"ooodev.char2.series.data_series.label.borders", "ooodev.char2.series.data_point.label.borders"}:
        from ooodev.format.inner.direct.chart2.series.data_labels.borders.line_properties import (
            LineProperties as LineProperties37,
        )

        return LineProperties37

    if name == "ooodev.char2.title":
        from ooodev.format.inner.direct.chart2.title.borders.line_properties import LineProperties as LineProperties38

        return LineProperties38

    raise ValueError(f"Invalid name: {name}")


def font_highlight_factory(name: str) -> Type[WriteCharFontHighlightT]:
    if name == "ooodev.write.char":
        from ooodev.format.inner.direct.write.char.highlight.highlight import Highlight as Highlight1

        return Highlight1

    raise ValueError(f"Invalid name: {name}")


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.write.char.font.font_only import FontOnly as FontOnly1
    from ooodev.format.inner.direct.chart2.axis.font.font_only import FontOnly as FontOnly2
    from ooodev.format.inner.direct.chart2.title.font.font_only import FontOnly as FontOnly3
    from ooodev.format.inner.direct.chart2.series.data_labels.font.font_only import FontOnly as FontOnly4
    from ooodev.format.inner.direct.chart2.legend.font.font_only import FontOnly as FontOnly5
    from ooodev.format.inner.direct.calc.char.font.font_only import FontOnly as FontOnly6
    from ooodev.format.inner.direct.write.char.font.font import Font as Font1
    from ooodev.format.inner.direct.general_style.text.font import Font as Font2
    from ooodev.format.inner.direct.write.char.font.font_effects import FontEffects as FontEffects1
    from ooodev.format.inner.direct.chart2.axis.font.font_effects import FontEffects as FontEffects2
    from ooodev.format.inner.direct.chart2.title.font.font_effects import FontEffects as FontEffects3
    from ooodev.format.inner.direct.chart2.series.data_labels.font.font_effects import FontEffects as FontEffects4
    from ooodev.format.inner.direct.chart2.legend.font.font_effects import FontEffects as FontEffects5
    from ooodev.format.inner.direct.calc.char.font.font_effects import FontEffects as FontEffects6
    from ooodev.format.inner.direct.write.char.font.font_position import FontPosition as FontPosition1
    from ooodev.format.inner.direct.calc.alignment.text_align import TextAlign as TextAlign1
    from ooodev.format.inner.direct.calc.alignment.text_orientation import TextOrientation as TextOrientation1
    from ooodev.format.inner.direct.calc.alignment.properties import Properties as Properties1
    from ooodev.format.inner.direct.calc.border.borders import Borders as Borders1
    from ooodev.format.inner.direct.calc.numbers.numbers import Numbers as Numbers1
    from ooodev.format.inner.direct.chart2.chart.numbers.numbers import Numbers as Numbers2
    from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.number_format import (
        NumberFormat as NumberFormat1,
    )
    from ooodev.format.inner.direct.chart2.axis.numbers.numbers import Numbers as NumberFormat2
    from ooodev.format.inner.direct.chart2.chart.area.color import Color as Color1
    from ooodev.format.inner.direct.chart2.legend.area.color import Color as Color2
    from ooodev.format.inner.direct.chart2.wall.area.color import Color as Color3
    from ooodev.format.inner.direct.chart2.series.data_series.area.color import Color as Color4
    from ooodev.format.inner.direct.chart2.title.area.color import Color as Color5
    from ooodev.format.inner.direct.calc.background.color import Color as Color6
    from ooodev.format.inner.direct.write.table.background.color import Color as Color7
    from ooodev.format.draw.direct.transparency.transparency import Transparency as Transparency1
    from ooodev.format.inner.direct.write.fill.transparent.transparency import Transparency as Transparency2
    from ooodev.format.writer.direct.shape.transparency.transparency import Transparency as Transparency3
    from ooodev.format.inner.direct.chart2.chart.transparent.transparency import Transparency as Transparency4
    from ooodev.format.inner.direct.chart2.legend.transparent.transparency import Transparency as Transparency5
    from ooodev.format.inner.direct.chart2.wall.transparent.transparency import Transparency as Transparency6
    from ooodev.format.inner.direct.chart2.series.data_series.transparent.transparency import (
        Transparency as Transparency7,
    )
    from ooodev.format.draw.direct.transparency.gradient import Gradient as Gradient1
    from ooodev.format.writer.direct.shape.transparency.gradient import Gradient as Gradient2
    from ooodev.format.inner.direct.chart2.chart.transparent.gradient import Gradient as Gradient3
    from ooodev.format.inner.direct.chart2.legend.transparent.gradient import Gradient as Gradient4
    from ooodev.format.inner.direct.chart2.wall.transparent.gradient import Gradient as Gradient5
    from ooodev.format.inner.direct.chart2.series.data_series.transparent.gradient import Gradient as Gradient6
    from ooodev.format.inner.direct.chart2.chart.area.gradient import Gradient as Gradient20
    from ooodev.format.inner.direct.chart2.legend.area.gradient import Gradient as Gradient21
    from ooodev.format.inner.direct.chart2.wall.area.gradient import Gradient as Gradient22
    from ooodev.format.inner.direct.chart2.series.data_series.area.gradient import Gradient as Gradient23
    from ooodev.format.inner.direct.chart2.title.area.gradient import Gradient as Gradient24
    from ooodev.format.inner.direct.chart2.chart.area.img import Img as Img20
    from ooodev.format.inner.direct.chart2.legend.area.img import Img as Img21
    from ooodev.format.inner.direct.chart2.wall.area.img import Img as Img22
    from ooodev.format.inner.direct.chart2.series.data_series.area.img import Img as Img23
    from ooodev.format.inner.direct.chart2.title.area.img import Img as Img24
    from ooodev.format.inner.direct.chart2.chart.area.pattern import Pattern as Pattern20
    from ooodev.format.inner.direct.chart2.legend.area.pattern import Pattern as Pattern21
    from ooodev.format.inner.direct.chart2.wall.area.pattern import Pattern as Pattern22
    from ooodev.format.inner.direct.chart2.series.data_series.area.pattern import Pattern as Pattern23
    from ooodev.format.inner.direct.chart2.title.area.pattern import Pattern as Pattern24
    from ooodev.format.inner.direct.chart2.chart.area.hatch import Hatch as Hatch20
    from ooodev.format.inner.direct.chart2.legend.area.hatch import Hatch as Hatch21
    from ooodev.format.inner.direct.chart2.wall.area.hatch import Hatch as Hatch22
    from ooodev.format.inner.direct.chart2.series.data_series.area.hatch import Hatch as Hatch23
    from ooodev.format.inner.direct.chart2.title.area.hatch import Hatch as Hatch24
    from ooodev.format.inner.direct.chart2.position_size.position import Position as Position20
    from ooodev.format.inner.direct.chart2.title.position_size.position import Position as Position21
    from ooodev.format.inner.direct.chart2.position_size.size import Size as Size20
    from ooodev.format.inner.direct.chart2.axis.positioning.axis_line import AxisLine as AxisLine20
    from ooodev.format.inner.direct.chart2.axis.positioning.interval_marks import IntervalMarks as IntervalMarks20
    from ooodev.format.inner.direct.chart2.axis.positioning.label_position import LabelPosition as LabelPosition20
    from ooodev.format.inner.direct.chart2.axis.positioning.position_axis import PositionAxis as PositionAxis20
    from ooodev.format.draw.direct.position_size.position_size.position import Position as Position30
    from ooodev.format.draw.direct.position_size.position_size.size import Size as Size30
    from ooodev.format.draw.direct.position_size.position_size.protect import Protect as Protect30
    from ooodev.format.inner.direct.chart2.chart.borders.line_properties import LineProperties as LineProperties30
    from ooodev.format.inner.direct.chart2.chart.borders.line_properties import LineProperties as LineProperties31
    from ooodev.format.inner.direct.chart2.axis.line.line_properties import LineProperties as LineProperties32
    from ooodev.format.inner.direct.chart2.grid.line_properties import LineProperties as LineProperties33
    from ooodev.format.inner.direct.chart2.legend.borders.line_properties import LineProperties as LineProperties34
    from ooodev.format.inner.direct.chart2.wall.borders.line_properties import LineProperties as LineProperties35
    from ooodev.format.inner.direct.chart2.series.data_series.borders.line_properties import (
        LineProperties as LineProperties36,
    )
    from ooodev.format.inner.direct.chart2.series.data_labels.borders.line_properties import (
        LineProperties as LineProperties37,
    )
    from ooodev.format.inner.direct.chart2.title.borders.line_properties import LineProperties as LineProperties38
    from ooodev.format.inner.direct.write.char.highlight.highlight import Highlight as Highlight1
