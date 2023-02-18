"""
Module for Fill Gradient Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar
import uno
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle
from ooo.dyn.awt.gradient import Gradient as UNOGradient

from .....events.args.cancel_event_args import CancelEventArgs
from .....events.args.key_val_cancel_args import KeyValCancelArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import color as mColor
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.data_type.angle import Angle as Angle
from .....utils.data_type.intensity import Intensity as Intensity
from .....utils.data_type.intensity_range import IntensityRange as IntensityRange
from .....utils.data_type.offset import Offset as Offset
from ....kind.format_kind import FormatKind
from ....style_base import StyleMulti
from ...structs.gradient_struct import GradientStruct
from .....events.format_named_event import FormatNamedEvent as FormatNamedEvent
from ...common.props.transparent_gradient_props import TransparentGradientProps


# from ooo.dyn.drawing.fill_style import FillStyle

# See Also:
# https://wiki.documentfoundation.org/Documentation/BASIC_Guide#Color_Gradient
# https://github.com/LibreOffice/core/blob/d57836db76fcf3133e6eb54d264c774911015e08/chart2/source/controller/itemsetwrapper/GraphicPropertyItemConverter.cxx
# https://github.com/LibreOffice/core/blob/cfb2a587bc59d2a0ff520dd09393f898506055d6/vcl/source/outdev/transparent.cxx
# https://github.com/LibreOffice/core/blob/7c3ea0abeff6e0cb9e2893cec8ed63025a274117/oox/source/export/drawingml.cxx

_TGradient = TypeVar(name="_TGradient", bound="Gradient")


class FillTransparentGrad(GradientStruct):
    def _supported_services(self) -> Tuple[str, ...]:
        if not self.prop_parent is None:
            return self.prop_parent._supported_services()
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.text.TextContent",
            "com.sun.star.style.ParagraphStyle",
            "com.sun.star.style.PageStyle",
        )

    def _get_property_name(self) -> str:
        try:
            return self._struct_property_name
        except AttributeError:
            cargs = CancelEventArgs(source=self)
            cargs.event_data = "FillTransparenceGradient"
            e_name = self._get_uniquie_event_name(FormatNamedEvent._STYLE_GETING_PROPERY_MISSING_NAME)
            self._events.trigger(event_name=e_name, event_args=cargs)
            if cargs.cancel:
                if cargs.handled:
                    return self._struct_property_name
                else:
                    raise mEx.CancelEventError(cargs)
            self._struct_property_name = cargs.event_data
        return self._struct_property_name

    def _container_get_service_name(self) -> str:
        return "com.sun.star.drawing.TransparencyGradientTable"

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        if not self.prop_parent is None:
            return self.prop_parent.prop_format_kind
        return FormatKind.PARA | FormatKind.TXT_CONTENT | FormatKind.FILL


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
            offset (offset, optional): Specifies the X-coordinate (start) and Y-coordinate (end), where the gradient begins.
                X is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT`` style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to 0.
            border (int, optional): Specifies percent of the total width where just the start color is used. Defaults to 0.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of the gradient. Defaults to ``IntensityRange(0, 0)``.
        """

        start_color = int(mColor.get_gray_rgb(grad_intensity.start))
        end_color = int(mColor.get_gray_rgb(grad_intensity.end))
        # start_color = 4144959
        # end_color = 16777215

        fs = FillTransparentGrad(
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
        # Fill Transparence is always zero when Gradient Tranparency is applied
        self._set(self._props.transparence, 0)

    # region Internal Methods
    def _set_fill_tp(self, fill_tp: FillTransparentGrad, name: str = "") -> None:
        fs = self._get_fill_tp(fill_tp, name)
        fs._prop_parent = self
        # add event listener to set the value returned by fs._get_property_name()
        fs.add_event_listener(FormatNamedEvent._STYLE_GETING_PROPERY_MISSING_NAME, _on_fill_get_property_name)
        self._set(self._props.name, self._name)
        self._set_style("fill_style", fs, *fs.get_attrs())

    def _get_fill_tp(self, fill_tp: FillTransparentGrad, name: str) -> FillTransparentGrad:
        # if the name passed in already exist in the TransparencyGradientTable Table then it is returned.
        # Otherwise the struc is added to the TransparencyGradientTable Table and then returned.
        # after struct is added to table all other subsequent call of this name will return
        # that struc from the Table. With the exception of auto_name which will force a new entry
        # into the Table each time.
        # see: https://github.com/LibreOffice/core/blob/d9e044f04ac11b76b9a3dac575f4e9155b67490e/chart2/source/tools/PropertyHelper.cxx#L212
        nc = self._container_get_inst()
        if name:
            struct = self._container_get_value(name, nc)  # raises value error if name is empty
            if not struct is None:
                self._name = name
                return FillTransparentGrad.from_gradient(struct)

        name = "Transparency "
        self._name = self._container_get_unique_el_name(name, nc)
        struct = self._container_get_value(self._name, nc)  # raises value error if name is empty
        if not struct is None:
            return FillTransparentGrad.from_gradient(struct)
        struct = fill_tp.get_uno_struct()
        self._container_add_value(name=self._name, obj=struct, allow_update=False, nc=nc)
        return FillTransparentGrad.from_gradient(self._container_get_value(self._name, nc))

    # endregion Internal Methods

    # region Overrides
    def copy(self: _TGradient) -> _TGradient:
        cp = super().copy()
        cp._name = self._name
        return cp

    def _container_get_service_name(self) -> str:
        return "com.sun.star.drawing.TransparencyGradientTable"

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
        if event_args.key == self._props.name:
            if event_args.value is None or event_args.value == "":
                # instruct Props.set to call set_default()
                event_args.default = True

        super().on_property_setting(event_args)

    # endregion Overrides

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
        nu = super(Gradient, cls).__new__(cls)
        nu.__init__()
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        grad_fill = cast(UNOGradient, mProps.Props.get(obj, nu._props.struct_prop))
        fill_gradient_name = cast(str, mProps.Props.get(obj, nu._props.name, ""))
        if grad_fill.Angle == 0:
            angle = 0
        else:
            angle = round(grad_fill.Angle / 10)
        return Gradient(
            style=grad_fill.Style,
            offset=Offset(grad_fill.XOffset, grad_fill.YOffset),
            angle=angle,
            border=Intensity(grad_fill.Border),
            grad_intensity=IntensityRange(grad_fill.StartIntensity, grad_fill.EndIntensity),
            transparency_name=fill_gradient_name,
        )

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.TXT_CONTENT | FormatKind.FILL

    @property
    def prop_inner(self) -> FillTransparentGrad:
        """Gets Fill Trasparent Gradient instance"""
        try:
            return self._direct_inner_fill_grad
        except AttributeError:
            self._direct_inner_fill_grad = cast(FillTransparentGrad, self._get_style_inst("fill_style"))
        return self._direct_inner_fill_grad

    @property
    def _props(self) -> TransparentGradientProps:
        try:
            return self._props_transparent_gradient
        except AttributeError:
            self._props_transparent_gradient = TransparentGradientProps(
                transparence="FillTransparence",
                name="FillTransparenceGradientName",
                struct_prop="FillTransparenceGradient",
            )
        return self._props_transparent_gradient

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
                angle=Angle(0),
                border=0,
                grad_intensity=IntensityRange(0, 0),
            )
            inst._is_default_inst = True
            Gradient._DEFAULT_INST = inst
        return Gradient._DEFAULT_INST


def _on_fill_get_property_name(source: Any, event_args: CancelEventArgs, *args, **kwargs) -> None:
    # this callback is only called once per instance of Gradient.
    # Sets the FillTransparentGrad._get_property_name() return value
    fg = cast(FillTransparentGrad, event_args.source)
    if not fg.prop_parent is None:
        grad = cast(Gradient, fg.prop_parent)
        event_args.event_data = grad._props.struct_prop
