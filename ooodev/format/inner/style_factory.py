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


def area_color_factory(name: str) -> Type[FillColorT]:
    if name == "ooodev.chart2.general":
        from ooodev.format.chart2.direct.general.area import Color

        return Color

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

    raise ValueError(f"Invalid name: {name}")
