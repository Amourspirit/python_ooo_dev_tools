"""
Module for Paragraph Gradient Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, cast
import uno
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils import color as mColor
from .....utils.data_type.angle import Angle as Angle
from .....utils.data_type.intensity import Intensity as Intensity
from ....kind.format_kind import FormatKind
from ....style_base import StyleMulti
from ...structs.gradient_struct import GradientStruct
from .....events.args.key_val_cancel_args import KeyValCancelArgs

from com.sun.star.container import XNameContainer

from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle
from ooo.dyn.awt.gradient import Gradient as UNOGradient
from ooo.dyn.drawing.fill_style import FillStyle

# See Also:
# https://wiki.documentfoundation.org/Documentation/BASIC_Guide#Color_Gradient
# https://github.com/LibreOffice/core/blob/d57836db76fcf3133e6eb54d264c774911015e08/chart2/source/controller/itemsetwrapper/GraphicPropertyItemConverter.cxx
# https://github.com/LibreOffice/core/blob/cfb2a587bc59d2a0ff520dd09393f898506055d6/vcl/source/outdev/transparent.cxx
# https://github.com/LibreOffice/core/blob/7c3ea0abeff6e0cb9e2893cec8ed63025a274117/oox/source/export/drawingml.cxx


class FillTransparendGrad(GradientStruct):
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.drawing.FillProperties", "com.sun.star.text.TextContent")

    def _get_property_name(self) -> str:
        return "FillTransparenceGradient"

    def _container_get_service_name(self) -> str:
        return "com.sun.star.drawing.TransparencyGradientTable"

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
        x_offset: Intensity | int = 50,
        y_offset: Intensity | int = 50,
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        start_value: Intensity | int = 0,
        end_value: Intensity | int = 0,
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
            start_value (Intensity, int, optional): Specifies the gradient start value from ``0`` to ``100``.
            end_value (Intensity, int, optional): Specifies the gradient End value from ``0`` to ``100``.
        """
        if not isinstance(start_value, Intensity):
            start_value = Intensity(start_value)
        if not isinstance(end_value, Intensity):
            end_value = Intensity(end_value)

        start_color = int(mColor.get_gray_rgb(start_value.value))
        end_color = int(mColor.get_gray_rgb(end_value.value))
        # start_color = 4144959
        # end_color = 16777215

        fs = FillTransparendGrad(
            style=style,
            step_count=0,
            x_offset=x_offset,
            y_offset=y_offset,
            angle=angle,
            border=border,
            start_color=start_color,
            start_intensity=100,
            end_color=end_color,
            end_intensity=100,
        )

        super().__init__()
        # gradient

        # FillTransparenceGradientName "Transparency " will be expaned to Transparency 1 or Transparency 2 etc in on_property_setting() method.
        self._set("FillTransparenceGradientName", "Transparency ")
        self._set("FillStyle", FillStyle.GRADIENT)
        # self._set("FillGradientName", "gradient")
        # self._set("FillHatchName", "hatch")
        self._set_style("fill_style", fs, *fs.get_attrs())

    def _container_get_service_name(self) -> str:
        return "com.sun.star.drawing.TransparencyGradientTable"

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.text.TextContent",
        )

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    def on_property_setting(self, event_args: KeyValCancelArgs) -> None:
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        if event_args.key == "FillTransparenceGradientName":
            if event_args.value is None or event_args.value == "":
                # instruct Props.set to call set_default()
                event_args.default = True
            elif event_args.value == "Transparency ":
                # https://github.com/LibreOffice/core/blob/d9e044f04ac11b76b9a3dac575f4e9155b67490e/chart2/source/tools/PropertyHelper.cxx#L212
                # get unique name
                nc = self._container_get_inst()
                name = self._container_get_unique_el_name(event_args.value, nc)

                # get the Gradient and add value to container
                # these are necessary steps to have gradient applied.
                grad = cast(FillTransparendGrad, self._get_style("fill_style")[0]).get_gradient()
                self._container_add_value(name=name, obj=grad, nc=nc)

                event_args.value = name

                # add to container.

        super().on_property_setting(event_args)

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

        gs = FillTransparendGrad()
        gs_prop_name = gs._get_property_name()

        grad_fill = cast(UNOGradient, mProps.Props.get(obj, gs_prop_name))
        gs = FillTransparendGrad.from_gradient(grad_fill)
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
            start_vaule=grad_fill.StartColor,
            start_intensity=grad_fill.StartIntensity,
            end_color=grad_fill.EndColor,
            end_intensity=grad_fill.EndIntensity,
            name=fill_gradient_name,
        )

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
                start_value=mColor.Color(0),
                end_value=mColor.Color(0),
            )
            Gradient._DEFAULT = inst
        return Gradient._DEFAULT

    @staticmethod
    def add_gradient_to_table():
        container = mLo.Lo.create_instance_msf(
            XNameContainer, "com.sun.star.drawing.TransparencyGradientTable", raise_err=True
        )
        names = container.getElementNames()
        for name in names:
            print(name)
