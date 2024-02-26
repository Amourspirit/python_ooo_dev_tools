"""
Module for Fill Gradient Color.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar, overload, TYPE_CHECKING
from ooo.dyn.awt.gradient_style import GradientStyle

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.transparent_gradient_props import TransparentGradientProps
from ooodev.format.inner.direct.structs.gradient_struct import GradientStruct
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.loader import lo as mLo
from ooodev.units.angle import Angle
from ooodev.utils import color as mColor
from ooodev.utils import props as mProps
from ooodev.utils.data_type.intensity import Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset

if TYPE_CHECKING:
    from ooo.dyn.awt.gradient import Gradient as UNOGradient

# endregion Import

# from ooo.dyn.drawing.fill_style import FillStyle

# See Also:
# https://wiki.documentfoundation.org/Documentation/BASIC_Guide#Color_Gradient
# https://github.com/LibreOffice/core/blob/d57836db76fcf3133e6eb54d264c774911015e08/chart2/source/controller/itemsetwrapper/GraphicPropertyItemConverter.cxx
# https://github.com/LibreOffice/core/blob/cfb2a587bc59d2a0ff520dd09393f898506055d6/vcl/source/outdev/transparent.cxx
# https://github.com/LibreOffice/core/blob/7c3ea0abeff6e0cb9e2893cec8ed63025a274117/oox/source/export/drawingml.cxx

_TGradient = TypeVar(name="_TGradient", bound="Gradient")


class Gradient(StyleMulti):
    """
    Fill Gradient Color

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_intensity: IntensityRange = IntensityRange(0, 0),
        **kwargs: Any,
    ) -> None:
        """
        Constructor

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (offset, optional): Specifies the X-coordinate (start) and Y-coordinate (end),
                where the gradient begins. X is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and
                ``RECT`` style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to ``0``.
            border (int, optional): Specifies percent of the total width where just the start color is used.
                Defaults to ``0``.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of
                the gradient. Defaults to ``IntensityRange(0, 0)``.

        Returns:
            None:
        """

        start_color = int(mColor.get_gray_rgb(grad_intensity.start))
        end_color = int(mColor.get_gray_rgb(grad_intensity.end))
        # start_color = 4144959
        # end_color = 16777215

        fs = self._get_inner_class(
            style=style,
            step_count=0,
            x_offset=offset.x,
            y_offset=offset.y,
            angle=angle,
            border=border,
            start_color=start_color,
            start_intensity=grad_intensity.start,
            end_color=end_color,
            end_intensity=grad_intensity.end,
        )

        super().__init__()
        # gradient
        self._set_fill_tp(fs, kwargs.get("transparency_name", ""))
        # self._set("FillStyle", FillStyle.SOLID)
        # Fill Transparence is always zero when Gradient Transparency is applied
        self._set(self._props.transparence, 0)

    # region Internal Methods
    def _get_inner_class(
        self,
        style: GradientStyle,
        step_count: int,
        x_offset: Intensity | int,
        y_offset: Intensity | int,
        angle: Angle | int,
        border: Intensity | int,
        start_color: int,
        start_intensity: Intensity | int,
        end_color: int,
        end_intensity: Intensity | int,
    ) -> GradientStruct:
        # pylint: disable=unexpected-keyword-arg
        return GradientStruct(
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
            _cattribs=self._get_inner_cattribs(),
        )

    def _get_inner_cattribs(self) -> dict:
        return {
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
            "_property_name": self._props.struct_prop,
        }

    def _set_fill_tp(self, fill_tp: GradientStruct, name: str = "") -> None:
        fs = self._get_fill_tp(fill_tp, name)

        self._set(self._props.name, self._name)
        self._set_style("fill_style", fs, *fs.get_attrs())

    def _container_get_default_name(self) -> str:
        return "Transparency"

    def _get_gradient_from_uno_struct(self, value: "UNOGradient", **kwargs) -> GradientStruct:
        return GradientStruct.from_uno_struct(value, **kwargs)

    def _get_fill_tp(self, fill_tp: GradientStruct, name: str) -> GradientStruct:
        # if the name passed in already exist in the TransparencyGradientTable Table then it is returned.
        # Otherwise, the struct is added to the TransparencyGradientTable Table and then returned.
        # after struct is added to table all other subsequent call of this name will return
        # that struct from the Table. Except auto_name which will force a new entry
        # into the Table each time.
        # see: https://github.com/LibreOffice/core/blob/d9e044f04ac11b76b9a3dac575f4e9155b67490e/chart2/source/tools/PropertyHelper.cxx#L212
        nc = self._container_get_inst()
        if name:
            struct = self._container_get_value(name, nc)  # raises value error if name is empty
            if struct is not None:
                self._name = name
                return self._get_gradient_from_uno_struct(value=struct, _cattribs=self._get_inner_cattribs())

        name = self._container_get_default_name()
        name = f"{name.strip()} "
        self._name = self._container_get_unique_el_name(name, nc)
        struct = self._container_get_value(self._name, nc)  # raises value error if name is empty
        if struct is not None:
            return self._get_gradient_from_uno_struct(value=struct, _cattribs=self._get_inner_cattribs())
        struct = fill_tp.get_uno_struct()
        self._container_add_value(name=self._name, obj=struct, allow_update=False, nc=nc)
        return self._get_gradient_from_uno_struct(
            value=self._container_get_value(self._name, nc), _cattribs=self._get_inner_cattribs()
        )

    # endregion Internal Methods

    # region Overrides
    # region copy()
    @overload
    def copy(self: _TGradient) -> _TGradient: ...

    @overload
    def copy(self: _TGradient, **kwargs) -> _TGradient: ...

    def copy(self: _TGradient, **kwargs) -> _TGradient:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._name = self._name
        return cp

    # endregion copy()

    def _container_get_service_name(self) -> str:
        return "com.sun.star.drawing.TransparencyGradientTable"

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
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # sourcery skip: merge-nested-ifs
        if event_args.key == self._props.name:
            if event_args.value is None or event_args.value == "":
                # instruct Props.set to call set_default()
                event_args.default = True

        super().on_property_setting(source, event_args)

    # endregion Overrides
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TGradient], obj: Any) -> _TGradient: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TGradient], obj: Any, **kwargs) -> _TGradient: ...

    @classmethod
    def from_obj(cls: Type[_TGradient], obj: Any, **kwargs) -> _TGradient:
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
        nu = cls(**kwargs)
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        grad_fill = cast("UNOGradient", mProps.Props.get(obj, nu._props.struct_prop))
        fill_gradient_name = cast(str, mProps.Props.get(obj, nu._props.name, ""))
        angle = 0 if grad_fill.Angle == 0 else round(grad_fill.Angle / 10)
        return cls(
            style=grad_fill.Style,
            offset=Offset(grad_fill.XOffset, grad_fill.YOffset),
            angle=angle,
            border=Intensity(grad_fill.Border),
            grad_intensity=IntensityRange(grad_fill.StartIntensity, grad_fill.EndIntensity),
            transparency_name=fill_gradient_name,
            **kwargs,
        )

    # endregion from_obj()
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA | FormatKind.TXT_CONTENT | FormatKind.FILL
        return self._format_kind_prop

    @property
    def prop_inner(self) -> GradientStruct:
        """Gets Fill Transparent Gradient instance"""
        try:
            return self._direct_inner_fill_grad
        except AttributeError:
            self._direct_inner_fill_grad = cast(GradientStruct, self._get_style_inst("fill_style"))
        return self._direct_inner_fill_grad

    @property
    def _props(self) -> TransparentGradientProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = TransparentGradientProps(
                transparence="FillTransparence",
                name="FillTransparenceGradientName",
                struct_prop="FillTransparenceGradient",
            )
        return self._props_internal_attributes

    @property
    def default(self: _TGradient) -> _TGradient:
        """Gets Gradient empty. Static Property."""
        try:
            return self._default_inst
        except AttributeError:
            inst = self.__class__(
                style=GradientStyle.LINEAR,
                step_count=0,
                offset=Offset(0, 0),
                angle=Angle(0),
                border=0,
                grad_intensity=IntensityRange(0, 0),
                _cattribs=self._get_inner_cattribs(),
            )
            inst._is_default_inst = True
            self._default_inst = inst
        return self._default_inst
