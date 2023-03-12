"""
Module for Paragraph Gradient Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, Type, cast, TypeVar, overload

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
        fs = GradientStruct(
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
            _cattribs=self._get_gradient_struct_cattrib(),
        )
        super().__init__()
        self._name = name

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

    def _get_gradient_struct_cattrib(self) -> dict:
        return {
            "_property_name": "FillGradient",
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    def _get_fill_struct(self, fill_struct: GradientStruct | None, name: str, auto_name: bool) -> GradientStruct:
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
            return GradientStruct.from_uno_struct(grad, _cattribs=self._get_gradient_struct_cattrib())
        if fill_struct is None:
            raise ValueError(
                f'No Gradient could be found in container for "{name}". In this case a Gradient is required.'
            )
        self._container_add_value(name=self._name, obj=fill_struct.get_uno_struct(), allow_update=False, nc=nc)
        return GradientStruct.from_uno_struct(
            self._container_get_value(self._name, nc), _cattribs=self._get_gradient_struct_cattrib()
        )

    def _container_get_service_name(self) -> str:
        # https://github.com/LibreOffice/core/blob/d9e044f04ac11b76b9a3dac575f4e9155b67490e/chart2/source/tools/PropertyHelper.cxx#L229
        return "com.sun.star.drawing.GradientTable"

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.drawing.FillProperties",
                "com.sun.star.style.PageStyle",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextContent",
                "com.sun.star.text.TextEmbeddedObject",
                "com.sun.star.text.TextGraphicObject",
            )
        return self._supported_services_values

    def _on_modifing(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(source, event)

    # region copy()
    @overload
    def copy(self: _TGradient) -> _TGradient:
        ...

    @overload
    def copy(self: _TGradient, **kwargs) -> _TGradient:
        ...

    def copy(self: _TGradient, **kwargs) -> _TGradient:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._name = self._name
        return cp

    # endregion copy()
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TGradient], obj: object) -> _TGradient:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TGradient], obj: object, **kwargs) -> _TGradient:
        ...

    @classmethod
    def from_obj(cls: Type[_TGradient], obj: object, **kwargs) -> _TGradient:
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
        inst = cls(name="__constructor_default__", **kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        grad_fill = cast(UNOGradient, mProps.Props.get(obj, inst._props.grad_prop_name))
        gs = GradientStruct.from_uno_struct(grad_fill, _cattribs=inst._get_gradient_struct_cattrib())

        fill_gradient_name = cast(str, mProps.Props.get(obj, inst._props.name))

        inst._set(inst._props.step_count, grad_fill.StepCount)
        inst._set(inst._props.name, fill_gradient_name)
        inst._set_style("fill_style", gs, *gs.get_attrs())
        return inst

    # endregion from_obj()

    # region from_gradient()
    @classmethod
    def from_struct(cls: Type[_TGradient], struct: GradientStruct, name: str = "", **kwargs) -> _TGradient:
        """
        Gets instance from ``GradientStruct`` instance

        Args:
            struct (GradientStruct): Gradient Struct instance.
            name (str, optional): Name of Gradient.

        Returns:
            Gradient:
        """
        if name and PresetGradientKind.is_preset(name):
            return cls.from_preset(PresetGradientKind(name), **kwargs)
        if name:
            auto_name = False
        else:
            auto_name = True
        inst = cls(name="__constructor_default__", **kwargs)
        grad_fill = struct.get_uno_struct()
        gs = GradientStruct.from_uno_struct(grad_fill, _cattribs=inst._get_gradient_struct_cattrib())
        fill_struct = inst._get_fill_struct(fill_struct=gs, name=name, auto_name=auto_name)

        inst._set(inst._props.step_count, grad_fill.StepCount)
        inst._set(inst._props.name, inst._name)
        inst._set_style("fill_style", fill_struct, *fill_struct.get_attrs())
        return inst

    # endregion from_gradient()

    # region from_preset()
    @overload
    @classmethod
    def from_preset(cls: Type[_TGradient], preset: PresetGradientKind) -> _TGradient:
        ...

    @overload
    @classmethod
    def from_preset(cls: Type[_TGradient], preset: PresetGradientKind, **kwargs) -> _TGradient:
        ...

    @classmethod
    def from_preset(cls: Type[_TGradient], preset: PresetGradientKind, **kwargs) -> _TGradient:
        """
        Gets instance from preset.

        Args:
            preset (PresetKind): Preset.

        Returns:
            Gradient: Graident from a preset.
        """
        args = preset_gradient.get_preset(preset)
        args.update(kwargs)
        return cls(**args)

    # endregion from_preset()

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA | FormatKind.TXT_CONTENT
        return self._format_kind_prop

    @property
    def prop_inner(self) -> GradientStruct:
        """Gets Fill styles instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(GradientStruct, self._get_style_inst("fill_style"))
        return self._direct_inner

    @property
    def prop_name(self) -> str:
        """Gets Current gradient Name"""
        return self._name

    @property
    def _props(self) -> AreaGradientProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaGradientProps(
                style="FillStyle",
                step_count="FillGradientStepCount",
                name="FillGradientName",
                grad_prop_name="FillGradient",
            )
        return self._props_internal_attributes

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
