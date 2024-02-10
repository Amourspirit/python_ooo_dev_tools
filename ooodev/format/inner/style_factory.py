from __future__ import annotations
from typing import Any, TYPE_CHECKING, Type

if TYPE_CHECKING:
    from ..proto.font.font_effects_t import FontEffectsT
    from ..proto.font.font_only_t import FontOnlyT
    from ..proto.area.fill_color_t import FillColorT
    from ..proto.area.fill_gradient_t import FillGradientT
    from ..proto.area.chart2.chart_fill_gradient_t import ChartFillGradientT
    from ..proto.area.chart2.chart_fill_img_t import ChartFillImgT
    from ..proto.area.chart2.chart_fill_pattern_t import ChartFillPatternT
    from ..proto.area.chart2.chart_fill_hatch_t import ChartFillHatchT
    from ..proto.position_size.chart2.position_t import PositionT as Chart2PositionT
    from ..proto.position_size.chart2.size_t import SizeT as Chart2SizeT
    from ..proto.position_size.draw.size_t import SizeT as DrawSizeT
    from ..proto.position_size.draw.position_t import PositionT as DrawPositionT
    from ..proto.position_size.draw.protect_t import ProtectT as DrawProtectT
    from ..proto.borders.line_properties_t import LinePropertiesT as BorderLinePropertiesT
    from ..proto.area.transparency.transparency_t import TransparencyT as TransparencyTransparencyT
    from ..proto.area.transparency.gradient_t import GradientT as TransparencyGradientT
    from ..proto.chart2.axis.positioning.axis_line_t import AxisLineT as Chart2AxisLineT
    from ..proto.chart2.axis.positioning.interval_marks_t import IntervalMarksT as Chart2IntervalMarksT
    from ..proto.chart2.axis.positioning.label_position_t import LabelPositionT as Chart2LabelPositionT
    from ..proto.chart2.axis.positioning.position_axis_t import PositionAxisT as Chart2PositionAxisT
    from ..proto.calc.numbers.numbers_t import NumbersT as CalcNumbersT
    from ..proto.chart2.numbers.numbers_t import NumbersT as Chart2NumbersT
    from ..proto.chart2.series.data_labels.data_labels.numbers_t import NumbersT as Chart2SeriesDataLabelsNumbersT
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


def font_only_factory(name: str) -> Type[FontOnlyT]:
    if name == "ooodev.write.char":
        from ooodev.format.inner.direct.write.char.font.font_only import FontOnly

        return FontOnly

    if name == "ooodev.chart2.axis":
        from ooodev.format.inner.direct.chart2.axis.font.font_only import FontOnly

        return FontOnly

    if name == "ooodev.chart2.title":
        from ooodev.format.chart2.direct.title.font import FontOnly

        return FontOnly

    if name == "ooodev.chart2.series.data_labels":
        from ooodev.format.chart2.direct.series.data_labels.font import FontOnly

        return FontOnly

    if name == "ooodev.chart2.legend":
        from ooodev.format.chart2.direct.legend.font import FontOnly

        return FontOnly

    raise ValueError(f"Invalid name: {name}")


def font_effects_factory(name: str) -> Type[FontEffectsT]:
    if name == "ooodev.write.char":
        from ooodev.format.inner.direct.write.char.font.font_effects import FontEffects

        return FontEffects
    if name == "ooodev.chart2.axis":
        from ooodev.format.inner.direct.chart2.axis.font.font_effects import FontEffects

        return FontEffects

    if name == "ooodev.chart2.title":
        from ooodev.format.chart2.direct.title.font import FontEffects

        return FontEffects

    if name == "ooodev.chart2.series.data_labels":
        from ooodev.format.chart2.direct.series.data_labels.font import FontEffects

        return FontEffects

    if name == "ooodev.chart2.legend":
        from ooodev.format.chart2.direct.legend.font import FontEffects

        return FontEffects

    raise ValueError(f"Invalid name: {name}")


# region numbers format
# region general numbers format
def numbers_numbers_factory(name: str) -> Type[CalcNumbersT]:
    if name == "ooodev.number.numbers":
        from ooodev.format.inner.direct.calc.numbers.numbers import Numbers

        return Numbers

    raise ValueError(f"Invalid name: {name}")


# endregion general numbers format


# region chart2 numbers format
def chart2_numbers_numbers_factory(name: str) -> Type[Chart2NumbersT]:
    if name == "ooodev.chart2.number.numbers":
        from ooodev.format.inner.direct.chart2.chart.numbers.numbers import Numbers

        return Numbers

    raise ValueError(f"Invalid name: {name}")


def chart2_series_data_labels_numbers_factory(name: str) -> Type[Chart2SeriesDataLabelsNumbersT]:
    if name == "ooodev.chart2.series.data_labels.number.numbers":
        from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.number_format import NumberFormat

        return NumberFormat

    if name == "ooodev.chart2.axis.numbers.numbers":
        from ooodev.format.inner.direct.chart2.axis.numbers.numbers import Numbers

        return Numbers

    raise ValueError(f"Invalid name: {name}")


# endregion chart2 numbers format
# endregion numbers format


def area_color_factory(name: str) -> Type[FillColorT]:
    if name == "ooodev.chart2.general":
        from ooodev.format.chart2.direct.general.area import Color

        return Color

    raise ValueError(f"Invalid name: {name}")


def area_transparency_transparency_factory(name: str) -> Type[TransparencyTransparencyT]:
    if name == "ooodev.area.transparency":
        from ooodev.format.draw.direct.transparency.transparency import Transparency

        return Transparency
    if name == "ooodev.write.para.transparency":
        from ooodev.format.writer.direct.para.transparency import Transparency

        return Transparency

    if name == "ooodev.write.shape.area.transparency":
        from ooodev.format.writer.direct.shape.transparency import Transparency

        return Transparency

    if name == "ooodev.chart2.general":
        from ooodev.format.chart2.direct.general.transparency import Transparency

        return Transparency

    raise ValueError(f"Invalid name: {name}")


def area_transparency_gradient_factory(name: str) -> Type[TransparencyGradientT]:
    if name == "ooodev.area.transparency":
        from ooodev.format.draw.direct.transparency.gradient import Gradient

        return Gradient

    if name == "ooodev.write.shape.area.transparency":
        from ooodev.format.writer.direct.shape.transparency import Gradient

        return Gradient

    if name == "ooodev.chart2.general":
        from ooodev.format.chart2.direct.general.transparency import Gradient

        return Gradient

    raise ValueError(f"Invalid name: {name}")


# def area_gradient_factory(name: str) -> Type[FillGradientT]:
#     if name == "ooodev.chart2.general":
#         # from ooodev.format.chart2.direct.general.area import Gradient
#         from ooodev.format.inner.direct.chart2.chart.area.gradient import Gradient

#         return Gradient

#     raise ValueError(f"Invalid name: {name}")


def chart2_area_gradient_factory(name: str) -> Type[ChartFillGradientT]:
    if name == "ooodev.chart2.general":
        # from ooodev.format.chart2.direct.general.area import Gradient
        from ooodev.format.inner.direct.chart2.chart.area.gradient import Gradient

        return Gradient

    raise ValueError(f"Invalid name: {name}")


def chart2_area_img_factory(name: str) -> Type[ChartFillImgT]:
    if name == "ooodev.chart2.general":
        # from ooodev.format.chart2.direct.general.area import Gradient
        from ooodev.format.inner.direct.chart2.chart.area.img import Img

        return Img

    raise ValueError(f"Invalid name: {name}")


def chart2_area_pattern_factory(name: str) -> Type[ChartFillPatternT]:
    if name == "ooodev.chart2.general":
        # from ooodev.format.chart2.direct.general.area import Gradient
        from ooodev.format.inner.direct.chart2.chart.area.pattern import Pattern

        return Pattern

    raise ValueError(f"Invalid name: {name}")


def chart2_area_hatch_factory(name: str) -> Type[ChartFillHatchT]:
    if name == "ooodev.chart2.general":
        # from ooodev.format.chart2.direct.general.area import Gradient
        from ooodev.format.inner.direct.chart2.chart.area.hatch import Hatch

        return Hatch

    raise ValueError(f"Invalid name: {name}")


def chart2_position_size_position_factory(name: str) -> Type[Chart2PositionT]:
    if name == "ooodev.chart2.position":
        from ooodev.format.inner.direct.chart2.position_size.position import Position

        return Position

    raise ValueError(f"Invalid name: {name}")


def chart2_position_size_size_factory(name: str) -> Type[Chart2SizeT]:
    if name == "ooodev.chart2.size":
        from ooodev.format.inner.direct.chart2.position_size.size import Size

        return Size

    raise ValueError(f"Invalid name: {name}")


# region Chart2 Axis
# region Chart2 Axis Positioning
def chart2_axis_pos_line_factory(name: str) -> Type[Chart2AxisLineT]:
    if name == "ooodev.chart2.axis.pos.line":
        from ooodev.format.inner.direct.chart2.axis.positioning.axis_line import AxisLine

        return AxisLine

    raise ValueError(f"Invalid name: {name}")


def chart2_axis_pos_interval_factory(name: str) -> Type[Chart2IntervalMarksT]:
    if name == "ooodev.chart2.axis.pos.interval_marks":
        from ooodev.format.inner.direct.chart2.axis.positioning.interval_marks import IntervalMarks

        return IntervalMarks

    raise ValueError(f"Invalid name: {name}")


def chart2_axis_pos_label_position_factory(name: str) -> Type[Chart2LabelPositionT]:
    if name == "ooodev.chart2.axis.pos.label_position":
        from ooodev.format.inner.direct.chart2.axis.positioning.label_position import LabelPosition

        return LabelPosition

    raise ValueError(f"Invalid name: {name}")


def chart2_axis_pos_position_axis_factory(name: str) -> Type[Chart2PositionAxisT]:
    if name == "ooodev.chart2.axis.pos.position":
        from ooodev.format.inner.direct.chart2.axis.positioning.position_axis import PositionAxis

        return PositionAxis

    raise ValueError(f"Invalid name: {name}")


# endregion Chart2 Axis Positioning

# endregion Chart2 Axis


def draw_position_size_position_factory(name: str) -> Type[DrawPositionT]:
    if name == "ooodev.draw.position":
        from ooodev.format.draw.direct.position_size.position_size.position import Position

        return Position

    raise ValueError(f"Invalid name: {name}")


def draw_position_size_size_factory(name: str) -> Type[DrawSizeT]:
    if name == "ooodev.draw.size":
        from ooodev.format.draw.direct.position_size.position_size.size import Size

        return Size

    raise ValueError(f"Invalid name: {name}")


def draw_position_size_protect_factory(name: str) -> Type[DrawProtectT]:
    if name == "ooodev.draw.protect":
        from ooodev.format.draw.direct.position_size.position_size.protect import Protect

        return Protect

    raise ValueError(f"Invalid name: {name}")


def draw_border_line_factory(name: str) -> Type[BorderLinePropertiesT]:
    if name == "ooodev.draw.line":
        from ooodev.format.inner.direct.chart2.chart.borders.line_properties import LineProperties

        return LineProperties

    if name == "ooodev.chart2.line":
        from ooodev.format.inner.direct.chart2.chart.borders.line_properties import LineProperties

        return LineProperties

    if name == "ooodev.chart2.axis.line":
        from ooodev.format.chart2.direct.axis.line import LineProperties

        return LineProperties

    raise ValueError(f"Invalid name: {name}")
