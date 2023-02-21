"""
Module for ``Gradient`` struct.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, Type, cast, overload, TypeVar, TYPE_CHECKING
import json

from ....exceptions import ex as mEx
from ....utils import props as mProps
from ....utils.color import Color, RGB
from ....utils.data_type.angle import Angle as Angle
from ....utils.data_type.intensity import Intensity as Intensity
from ...style_base import StyleBase
from ...kind.format_kind import FormatKind


import uno
from ooo.dyn.awt.gradient import Gradient
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle


# see Also:
# https://github.com/LibreOffice/core/blob/f725629a6241ec064770c28957f11d306c18f130/filter/source/msfilter/escherex.cxx

_TGradientStruct = TypeVar(name="_TGradientStruct", bound="GradientStruct")


class GradientStruct(StyleBase):
    """
    Represents UNO ``Gradient`` struct.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        step_count: int = 0,
        x_offset: Intensity | int = 50,
        y_offset: Intensity | int = 50,
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        start_color: Color = Color(0),
        start_intensity: Intensity | int = 100,
        end_color: Color = Color(16777215),
        end_intensity: Intensity | int = 100,
    ) -> None:
        """
        Constructor

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

        Raises:
            ValueError: If ``step_count`` is less then zero.

        Returns:
            None:
        """
        if not isinstance(angle, Angle):
            angle = Angle(angle)
        if not isinstance(end_intensity, Intensity):
            end_intensity = Intensity(end_intensity)
        if not isinstance(start_intensity, Intensity):
            start_intensity = Intensity(start_intensity)
        if not isinstance(x_offset, Intensity):
            x_offset = Intensity(x_offset)
        if not isinstance(y_offset, Intensity):
            y_offset = Intensity(y_offset)
        if not isinstance(border, Intensity):
            border = Intensity(border)
        # step_count must be between 3 and 256 when not automatic in paragraph gradient
        if step_count < 0:
            raise ValueError("step_count must be a postivie number")

        init_vals = {
            "Style": style,
            "StartColor": start_color,
            "EndColor": end_color,
            "Angle": angle.value * 10,
            "Border": border.value,
            "XOffset": x_offset.value,
            "YOffset": y_offset.value,
            "StartIntensity": start_intensity.value,
            "EndIntensity": end_intensity.value,
            "StepCount": step_count,
        }

        super().__init__(**init_vals)

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "FillGradient"
        return self._property_name

    def get_attrs(self) -> Tuple[str, ...]:
        return (self._get_property_name(),)

    def get_uno_struct(self) -> Gradient:
        """
        Gets UNO ``Gradient`` from instance.

        Returns:
            Gradient: ``Gradient`` instance
        """
        return Gradient(
            Style=self._get("Style"),
            StartColor=self._get("StartColor"),
            EndColor=self._get("EndColor"),
            Angle=self._get("Angle"),
            Border=self._get("Border"),
            XOffset=self._get("XOffset"),
            YOffset=self._get("YOffset"),
            StartIntensity=self._get("StartIntensity"),
            EndIntensity=self._get("EndIntensity"),
            StepCount=self._get("StepCount"),
        )

    def __eq__(self, oth: object) -> bool:
        obj2 = None
        if isinstance(oth, GradientStruct):
            obj2 = oth.get_uno_struct()
        if getattr(oth, "typeName", None) == "com.sun.star.awt.Gradient":
            obj2 = cast(Gradient, oth)
        if obj2:
            obj1 = self.get_uno_struct()
            return (
                obj1.Style == obj2.Style
                and obj1.StartColor == obj2.StartColor
                and obj1.EndColor == obj2.EndColor
                and obj1.Angle == obj2.Angle
                and obj1.Border == obj2.Border
                and obj1.XOffset == obj2.XOffset
                and obj1.YOffset == obj2.YOffset
                and obj1.StartIntensity == obj2.StartIntensity
                and obj1.EndIntensity == obj2.EndIntensity
                and obj1.StepCount == obj2.StepCount
            )
        return NotImplemented

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies tab properties to ``obj``

        Args:
            obj (object): UNO object.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        # override_dv
        if not mProps.Props.has(obj, self._get_property_name()):
            self._print_not_valid_obj("apply")
            return

        grad = self.get_uno_struct()
        props = {self._get_property_name(): grad}
        super().apply(obj=obj, override_dv=props)

    # endregion apply()

    # region JSON
    def get_json(self) -> str:
        """
        Get Gradient represented as a json string for use with dispatch commands.

        Returns:
            str: Json string.
        """
        # See Also: https://tinyurl.com/2p7o5tvt search for FillPageGradientJSON

        # It seems dispatch commands (at least Writer) do not seem to work.
        # possible dispatch commands are:
        # props = Props.make_props(FillPageGradientJSON=json_dat)
        # or
        # props = Props.make_props(FillGradientJSON=json_dat)
        #
        # Lo.dispatch_cmd("FillPageGradient", props)
        # or
        # Lo.dispatch_cmd("FillGradient", props)
        #
        # see: https://wiki.documentfoundation.org/Development/DispatchCommands search for .uno:FillGradient
        d = {
            "style": cast(uno.Enum, self._get("Style")).value,
            "startcolor": RGB.from_int(self._get("StartColor")).to_hex(),
            "endcolor": RGB.from_int(self._get("EndColor")).to_hex(),
            "angle": str(self._get("Angle")),
            "border": str(self._get("Border")),
            "x": str(self._get("XOffset")),
            "y": str(self._get("YOffset")),
            "intensstart": str(self._get("StartIntensity")),
            "intensend": str(self._get("EndIntensity")),
            "stepcount": str(self._get("StepCount")),
        }
        return json.dumps(d)

    # endregion JSON

    # region static methods
    # region from_gradient()
    @overload
    @classmethod
    def from_gradient(cls: Type[_TGradientStruct], value: Gradient) -> _TGradientStruct:
        ...

    @overload
    @classmethod
    def from_gradient(cls: Type[_TGradientStruct], value: Gradient, **kwargs) -> _TGradientStruct:
        ...

    @classmethod
    def from_gradient(cls: Type[_TGradientStruct], value: Gradient, **kwargs) -> _TGradientStruct:
        """
        Converts a ``Gradient`` instance to a ``GradinetStruct``

        Args:
            value (Gradient): UNO Gradient

        Returns:
            GradinetStruct: ``GradinetStruct`` set with ``Gradient`` properties
        """
        inst = cls(**kwargs)
        inst._set("Style", value.Style)
        inst._set("StartColor", value.StartColor)
        inst._set("EndColor", value.EndColor)
        inst._set("Angle", value.Angle)
        inst._set("Border", value.Border)
        inst._set("XOffset", value.XOffset)
        inst._set("YOffset", value.YOffset)
        inst._set("StartIntensity", value.StartIntensity)
        inst._set("EndIntensity", value.EndIntensity)
        inst._set("StepCount", value.StepCount)
        return inst

    # endregion from_gradient()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TGradientStruct], obj: object) -> _TGradientStruct:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TGradientStruct], obj: object, **kwargs) -> _TGradientStruct:
        ...

    @classmethod
    def from_obj(cls: Type[_TGradientStruct], obj: object, **kwargs) -> _TGradientStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            GradientStruct: ``GradientStruct`` instance that represents ``obj`` gradient properties.
        """
        # this nu is only used to get Property Name
        nu = cls(**kwargs)
        prop_name = nu._get_property_name()

        try:
            grad = cast(Gradient, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")

        return cls.from_gradient(grad, **kwargs)

    # endregion from_obj()

    # endregion static methods

    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA | FormatKind.TXT_CONTENT
        return self._format_kind_prop

    @property
    def prop_style(self) -> GradientStyle:
        """Gets/Sets the style of the gradient."""
        return self._get("Style")

    @prop_style.setter
    def prop_style(self, value: GradientStyle):
        self._set("Style", value)

    @property
    def prop_step_count(self) -> int:
        """Gets/Sets the number of steps of change color."""
        return self._get("StepCount")

    @prop_step_count.setter
    def prop_step_count(self, value: int):
        self._set("StepCount", value)

    @property
    def prop_x_offset(self) -> Intensity:
        """Gets/Sets the X-coordinate, where the gradient begins."""
        pv = cast(int, self._get("XOffset"))
        return Intensity(pv)

    @prop_x_offset.setter
    def prop_x_offset(self, value: Intensity | int):
        if not isinstance(value, Intensity):
            value = Intensity(value)
        self._set("XOffset", value.value)

    @property
    def prop_y_offset(self) -> Intensity:
        """Gets/Sets the Y-coordinate, where the gradient begins."""
        pv = cast(int, self._get("YOffset"))
        return Intensity(pv)

    @prop_y_offset.setter
    def prop_y_offset(self, value: Intensity | int):
        if not isinstance(value, Intensity):
            value = Intensity(value)
        self._set("YOffset", value.value)

    @property
    def prop_angle(self) -> Angle:
        """Gets/Sets angle of the gradient."""
        pv = cast(int, self._get("Angle"))
        if pv == 0:
            return Angle(0)
        return Angle(round(pv / 10))

    @prop_angle.setter
    def prop_angle(self, value: Angle | int):
        if not isinstance(value, Angle):
            value = Angle(value)
        self._set("Angle", value.value * 10)

    @property
    def prop_border(self) -> Intensity:
        """Gets/Sets percent of the total width where just the start color is used."""
        pv = cast(int, self._get("Border"))
        return Intensity(pv)

    @prop_border.setter
    def prop_border(self, value: Intensity | int):
        if not isinstance(value, Intensity):
            value = Intensity(value)
        self._set("Border", value.value)

    @property
    def prop_start_color(self) -> Color:
        """Gets/Sets the color at the start point of the gradient."""
        return self._get("StartColor")

    @prop_start_color.setter
    def prop_start_color(self, value: Color):
        self._set("StartColor", value)

    @property
    def prop_start_intensity(self) -> Intensity:
        """Gets/Sets the intensity at the start point of the gradient."""
        pv = cast(int, self._get("StartIntensity"))
        return Intensity(pv)

    @prop_start_intensity.setter
    def prop_start_intensity(self, value: Intensity | int):
        if not isinstance(value, Intensity):
            value = Intensity(value)
        self._set("StartIntensity", value.value)

    @property
    def prop_end_color(self) -> Color:
        """Gets/Sets the color at the end point of the gradient."""
        return self._get("EndColor")

    @prop_end_color.setter
    def prop_end_color(self, value: Color):
        self._set("EndColor", value)

    @property
    def prop_end_intensity(self) -> Intensity:
        """Gets/Sets the intensity at the end point of the gradient."""
        pv = cast(int, self._get("EndIntensity"))
        return Intensity(pv)

    @prop_end_intensity.setter
    def prop_end_intensity(self, value: Intensity | int):
        if not isinstance(value, Intensity):
            value = Intensity(value)
        self._set("EndIntensity", value.value)

    # endregion properties
