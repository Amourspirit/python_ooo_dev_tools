"""
Module for Paragraph Fill Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload

from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from .....exceptions import ex as mEx
from .....utils.color import Color
from .....utils.data_type.intensity import Intensity as Intensity
from .....utils.data_type.angle import Angle as Angle
from ....kind.format_kind import FormatKind
from ....style_base import StyleMulti
from ...structs.gradient_struct import GradinetStruct
from ....preset import preset_gradient

from ....preset.preset_gradient import PresetKind as PresetKind

from ooo.dyn.drawing.fill_style import FillStyle
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle


class FillStyleStruct(GradinetStruct):
    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.text.TextContent",
        )

    def _get_property_name(self) -> str:
        return "FillGradient"

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.FILL


class Gradient(StyleMulti):
    """
    Paragraph Drop Caps

    Warning:
        This class uses dispatch commands and is not suitable for use in headless mode.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    def __init__(
        self,
        style: GradientStyle = GradientStyle.LINEAR,
        step_count: int = 0,
        x_offset: Intensity | int = 50,
        y_offset: Intensity | int = 50,
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        start_color: Color = 0,
        start_intensity: Intensity | int = 100,
        end_color: Color = 16777215,
        end_intensity: Intensity | int = 100,
        name: str = "",
    ) -> None:
        fs = FillStyleStruct(
            style=style,
            step_count=step_count,
            x_offset=x_offset,
            y_offset=y_offset,
            angle=angle,
            border=border,
            start_color=start_color,
            start_intensity=start_intensity,
            end_color=end_color,
            end_intensity=end_intensity,
        )
        super().__init__()
        # gradient

        self._set("FillStyle", FillStyle.GRADIENT)
        self._set("FillGradientStepCount", step_count)
        self._set("FillGradientName", self._get_gradient_name(style, name))
        self._set_style("fill_style", fs, *fs.get_attrs())

    def _get_gradient_name(self, style: GradientStyle, name: str) -> str:
        if name:
            return name
        if style == GradientStyle.AXIAL:
            return "Gradient 2"
        elif style == GradientStyle.ELLIPTICAL:
            return "Gradient 3"
        elif style == GradientStyle.SQUARE:
            # Square is quadratic
            return "Gradient 8"
        elif style == GradientStyle.RECT:
            # Rect is Square
            return "Gradient 7"
        else:
            return "gradient"

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.text.TextContent",
        )

    @staticmethod
    def from_preset(preset: PresetKind) -> Gradient:
        args = preset_gradient.get_preset(preset)
        return Gradient(**args)

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.FILL

    @static_prop
    def default() -> Gradient:  # type: ignore[misc]
        """Gets Gradient empty. Static Property."""
        if Gradient._DEFAULT is None:
            inst = Gradient(
                style=GradientStyle.LINEAR,
                step_count=0,
                x_offset=0,
                y_offset=0,
                angle=0,
                border=0,
                start_color=Color(0),
                start_intensity=100,
                end_color=Color(16777215),
                end_intensity=100,
            )
            inst._set("FillStyle", FillStyle.NONE)
            inst._set("FillGradientName", "")
            Gradient._DEFAULT = inst
        return Gradient._DEFAULT
