"""
Module for Paragraph Gradient Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, Type, cast, TypeVar

import uno
from ooo.dyn.drawing.fill_style import FillStyle
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle
from ooo.dyn.awt.gradient import Gradient as UNOGradient

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import props as mProps
from .....utils.color import Color
from .....utils.data_type.angle import Angle as Angle
from .....utils.data_type.color_range import ColorRange as ColorRange
from .....utils.data_type.intensity import Intensity as Intensity
from .....utils.data_type.intensity_range import IntensityRange as IntensityRange
from .....utils.data_type.offset import Offset as Offset
from ....kind.format_kind import FormatKind
from ....preset import preset_gradient
from ....preset.preset_gradient import PresetGradientKind as PresetGradientKind
from ....style_base import StyleMulti
from ...common.props.area_gradient_props import AreaGradientProps
from ...structs.gradient_struct import GradientStruct


_TGradient = TypeVar(name="_TGradient", bound="Gradient")


class FillStyleStruct(GradientStruct):
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_struct_services
        except AttributeError:
            self._supported_struct_services = (
                "com.sun.star.drawing.FillProperties",
                "com.sun.star.text.TextContent",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.style.PageStyle",
            )
        return self._supported_struct_services

    def _get_property_name(self) -> str:
        try:
            return self._struct_property_name
        except AttributeError:
            self._struct_property_name = "FillGradient"
        return self._struct_property_name

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.TXT_CONTENT | FormatKind.FILL


class Gradient(StyleMulti):
    """
    Paragraph Gradient Color


    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        step_count: int = 0,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_color: ColorRange = ColorRange(Color(0), Color(16777215)),
        grad_intensity: IntensityRange = IntensityRange(100, 100),
        name: str = "",
    ) -> None:
        """
        Constructor

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (Offset, int, optional): Specifies the X and Y coordinate, where the gradient begins.
                 X is is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT`` style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to 0.
            border (int, optional): Specifies percent of the total width where just the start color is used. Defaults to 0.
            grad_color (ColorRange, optional): Specifies the color at the start point and stop point of the gradient. Defaults to ``ColorRange(Color(0), Color(16777215))``.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of the gradient. Defaults to ``IntensityRange(100, 100)``.
            name (str, optional): Specifies the Fill Gradient Name.
        """
        fs = FillStyleStruct(
            style=style,
            step_count=step_count,
            x_offset=offset.x,
            y_offset=offset.y,
            angle=angle,
            border=border,
            start_color=grad_color.start,
            start_intensity=grad_intensity.start,
            end_color=grad_color.end,
            end_intensity=grad_intensity.end,
        )
        fs._struct_property_name = self._props.grad_prop_name
        fs._supported_struct_services = self._supported_services()
        super().__init__()

        self._set(self._props.style, FillStyle.GRADIENT)
        self._set(self._props.step_count, step_count)
        if name == "__constructor_default__":
            # __constructor_default__ means the values will never be used.
            self._set(self._props.name, "Gradient 9999")
            self._set_style("fill_style", fs, *fs.get_attrs())
        else:
            fill_struct = self._get_fill_struct(fill_struct=fs, name=name, auto_name=False)
            self._set(self._props.name, self._name)
            self._set_style("fill_style", fill_struct, *fill_struct.get_attrs())

    def _get_fill_struct(self, fill_struct: FillStyleStruct | None, name: str, auto_name: bool) -> FillStyleStruct:
        # if the name passed in already exist in the Gradient Table then it is returned.
        # Otherwise the Gradient is added to the Gradient Table and then returned.
        # after Gradient is added to table all other subsequent call of this name will return
        # that Gradient from the Table. With the exception of auto_name which will force a new entry
        # into the Table each time.
        self._name = name
        if name:
            if PresetGradientKind.is_preset(name):
                return fill_struct
        else:
            auto_name = True
            name = "Gradient"
        nc = self._container_get_inst()
        if auto_name:
            name = name.rstrip() + " "  # add a space after name before getting unique name
            self._name = self._container_get_unique_el_name(name, nc)

        grad = self._container_get_value(self._name, nc)  # raises value error if name is empty
        if not grad is None:
            return FillStyleStruct.from_gradient(grad)
        if fill_struct is None:
            raise ValueError(
                f'No Gradient could be found in container for "{name}". In this case a Gradient is required.'
            )
        self._container_add_value(name=self._name, obj=fill_struct.get_uno_struct(), allow_update=False, nc=nc)
        return FillStyleStruct.from_gradient(self._container_get_value(self._name, nc))

    def _container_get_service_name(self) -> str:
        # https://github.com/LibreOffice/core/blob/d9e044f04ac11b76b9a3dac575f4e9155b67490e/chart2/source/tools/PropertyHelper.cxx#L229
        return "com.sun.star.drawing.GradientTable"

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.text.TextContent",
            "com.sun.star.style.ParagraphStyle",
            "com.sun.star.style.PageStyle",
        )

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    @classmethod
    def from_obj(cls: Type[_TGradient], obj: object) -> _TGradient:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Gradient: Instance that represents Gradient color.
        """
        # this nu is only used to get Property Name
        inst = super(Gradient, cls).__new__(cls)
        inst.__init__(name="__constructor_default__")
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        grad_fill = cast(UNOGradient, mProps.Props.get(obj, inst._props.grad_prop_name))
        gs = FillStyleStruct.from_gradient(grad_fill)
        gs._struct_property_name = inst._props.grad_prop_name
        gs._supported_struct_services = inst._supported_services()

        fill_gradient_name = cast(str, mProps.Props.get(obj, inst._props.name))

        inst._set(inst._props.step_count, grad_fill.StepCount)
        inst._set(inst._props.name, fill_gradient_name)
        inst._set_style("fill_style", gs, *gs.get_attrs())
        return inst

    @classmethod
    def from_preset(cls: Type[_TGradient], preset: PresetGradientKind) -> _TGradient:
        """
        Gets instance from preset.

        Args:
            preset (PresetKind): Preset.

        Returns:
            Gradient: Graident from a preset.
        """
        args = preset_gradient.get_preset(preset)
        inst = super(Gradient, cls).__new__(cls)
        inst.__init__(**args)
        return inst

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.TXT_CONTENT

    @property
    def prop_inner(self) -> FillStyleStruct:
        """Gets Fill styles instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(FillStyleStruct, self._get_style_inst("fill_style"))
        return self._direct_inner

    @property
    def _props(self) -> AreaGradientProps:
        try:
            return self._props_gradient
        except AttributeError:
            self._props_gradient = AreaGradientProps(
                style="FillStyle",
                step_count="FillGradientStepCount",
                name="FillGradientName",
                grad_prop_name="FillGradient",
            )
        return self._props_gradient

    @static_prop
    def default() -> Gradient:  # type: ignore[misc]
        """Gets Gradient empty. Static Property."""
        try:
            return Gradient._DEFAULT_INST
        except AttributeError:
            inst = Gradient(
                style=GradientStyle.LINEAR,
                step_count=0,
                offset=Offset(0, 0),
                angle=0,
                border=0,
                grad_color=ColorRange(0, 16777215),
                grad_intensity=IntensityRange(100, 100),
            )
            inst._set(inst._props.style, FillStyle.NONE)
            inst._set(inst._props.name, "")
            inst._is_default_inst = True
            Gradient._DEFAULT_INST = inst
        return Gradient._DEFAULT_INST
