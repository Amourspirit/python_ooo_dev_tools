"""
Module for Paragraph Gradient Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast

from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import props as mProps
from .....utils.color import Color
from .....utils.data_type.angle import Angle as Angle
from .....utils.data_type.intensity import Intensity as Intensity
from ....kind.format_kind import FormatKind
from ....preset import preset_gradient
from ....preset.preset_gradient import PresetGradientKind as PresetGradientKind
from ....style_base import StyleMulti
from ...structs.gradient_struct import GradientStruct

from ooo.dyn.drawing.fill_style import FillStyle
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle
from ooo.dyn.awt.gradient import Gradient as UNOGradient


class FillStyleStruct(GradientStruct):
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.drawing.FillProperties", "com.sun.star.text.TextContent")

    def _get_property_name(self) -> str:
        return "FillGradient"

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.TXT_CONTENT | FormatKind.FILL


class Gradient(StyleMulti):
    """
    Paragraph Gradient Color


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
        """
        _summary_

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            x_offset (Intensity, int, optional): Specifies the X-coordinate, where the gradient begins.
                This is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT`` style gradients. Defaults to ``50``.
            y_offset (Intensity, int, optional): Specifies the Y-coordinate, where the gradient begins.
                See: ``x_offset``. Defaults to ``50``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to 0.
            border (int, optional): Specifies percent of the total width where just the start color is used. Defaults to 0.
            start_color (Color, optional): Specifies the color at the start point of the gradient. Defaults to ``Color(0)``.
            start_intensity (Intensity, int, optional): Specifies the intensity at the start point of the gradient. Defaults to ``100``.
            end_color (Color, optional): Specifies the color at the end point of the gradient. Defaults to ``Color(16777215)``.
            end_intensity (Intensity, int, optional): Specifies the intensity at the end point of the gradient. Defaults to ``100``.
            name (str, optional): Specifies the Fill Gradient Name.
        """
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

    @classmethod
    def from_obj(cls, obj: object) -> Gradient:
        """
        Gets instance from object

        Args:
            obj (object): Object that implements ``com.sun.star.drawing.FillProperties`` service

        Returns:
            Gradient: Instance that represents Gradient color.
        """
        # this nu is only used to get Property Name
        nu = super(Gradient, cls).__new__(cls)
        nu.__init__()
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError("obj is not supported")

        gs = FillStyleStruct()
        gs_prop_name = gs._get_property_name()

        grad_fill = cast(UNOGradient, mProps.Props.get(obj, gs_prop_name))
        gs = FillStyleStruct.from_gradient(grad_fill)
        fill_gradient_name = cast(str, mProps.Props.get(obj, "FillGradientName"))
        if grad_fill.Angle == 0:
            angle = 0
        else:
            angle = round(grad_fill.Angle / 10)
        return Gradient(
            style=grad_fill.Style,
            step_count=grad_fill.StepCount,
            x_offset=grad_fill.XOffset,
            y_offset=grad_fill.YOffset,
            angle=angle,
            border=grad_fill.Border,
            start_color=grad_fill.StartColor,
            start_intensity=grad_fill.StartIntensity,
            end_color=grad_fill.EndColor,
            end_intensity=grad_fill.EndIntensity,
            name=fill_gradient_name,
        )

    @staticmethod
    def from_preset(preset: PresetGradientKind) -> Gradient:
        """
        Gets instance from preset

        Args:
            preset (PresetKind): Preset

        Returns:
            Gradient: Graident from a preset.
        """
        args = preset_gradient.get_preset(preset)
        return Gradient(**args)

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.TXT_CONTENT

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
