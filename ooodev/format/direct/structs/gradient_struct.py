"""
Module for ``Gradient`` struct.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload
import json
from enum import Enum

from ....events.event_singleton import _Events
from ....exceptions import ex as mEx
from ....utils import props as mProps
from ....utils.color import Color, RGB
from ....utils.data_type.angle import Angle as Angle
from ....utils.data_type.intensity import Intensity as Intensity
from ....utils.type_var import T
from ...style_base import StyleBase, EventArgs, CancelEventArgs, FormatNamedEvent
from ...kind.format_kind import FormatKind


import uno
from ooo.dyn.awt.gradient import Gradient
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle


class GradinetStruct(StyleBase):
    """
    Represents UNO ``Gradient`` struct.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
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
            x_offset (Intensity | int, optional): Specifies the X-coordinate, where the gradient begins.
                This is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT`` style gradients. Defaults to ``50``.
            y_offset (Intensity | int, optional): Specifies the Y-coordinate, where the gradient begins.
                See: ``x_offset``. Defaults to ``50``.
            angle (Angle | int, optional): Specifies angle of the gradient. Defaults to 0.
            border (int, optional): Specifies percent of the total width where just the start color is used. Defaults to 0.
            start_color (Color, optional): Specifies the color at the start point of the gradient. Defaults to ``Color(0)``.
            start_intensity (Intensity | int, optional): Specifies the intensity at the start point of the gradient. Defaults to ``100``.
            end_color (Color, optional): Specifies the color at the end point of the gradient. Defaults to ``Color(16777215)``.
            end_intensity (Intensity | int, optional): Specifies the intensity at the end point of the gradient. Defaults to ``100``.

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
        return ()

    def _get_property_name(self) -> str:
        return "FillGradient"

    def copy(self: T) -> T:
        nu = super(GradinetStruct, self.__class__).__new__(self.__class__)
        nu.__init__()
        if self._dv:
            nu._update(self._dv)
        return nu

    def get_attrs(self) -> Tuple[str, ...]:
        return (self._get_property_name(),)

    def get_gradient(self) -> Gradient:
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
        if isinstance(oth, GradinetStruct):
            obj2 = oth.get_gradient()
        if getattr(oth, "typeName", None) == "com.sun.star.awt.Gradient":
            obj2 = oth
        if obj2:
            obj1 = self.get_gradient()
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

        If a DropCap instance with the same position is existing it is updated;
        Otherwise, a new DropCap is added.

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.
            keys (Dict[str, str], optional): Property key, value items that map properties.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        if not self._is_valid_obj(obj):
            # will not apply on this class but may apply on child classes
            self._print_not_valid_obj("apply")
            return

        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        self.on_applying(cargs)
        if cargs.cancel:
            return
        _Events().trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        grad = self.get_gradient()
        mProps.Props.set(obj, **{self._get_property_name(): grad})
        eargs = EventArgs.from_args(cargs)
        self.on_applied(eargs)
        _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

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
    @classmethod
    def from_gradient(cls, grad: Gradient) -> GradinetStruct:
        """
        Converts a ``Gradient`` instance to a ``GradinetStruct``

        Args:
            grad (Gradient): UNO Gradient

        Returns:
            GradinetStruct: ``GradinetStruct`` set with ``Gradient`` properties
        """
        inst = super(GradinetStruct, cls).__new__(cls)
        inst.__init__()
        inst._set("Style", grad.Style),
        inst._set("StartColor", grad.StartColor),
        inst._set("EndColor", grad.EndColor),
        inst._set("Angle", grad.Angle),
        inst._set("Border", grad.Border),
        inst._set("XOffset", grad.XOffset),
        inst._set("YOffset", grad.YOffset),
        inst._set("StartIntensity", grad.StartIntensity),
        inst._set("EndIntensity", grad.EndIntensity),
        inst._set("StepCount", grad.StepCount),
        return inst

    @classmethod
    def from_obj(cls, obj: object) -> GradinetStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            DropCap: ``DropCap`` instance that represents ``obj`` Drop cap format properties.
        """
        # this nu is only used to get Property Name
        nu = super(Gradient, cls).__new__(cls)
        nu.__init__()
        prop_name = nu._get_property_name()

        try:
            grad = cast(Gradient, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")

        return cls.from_gradient(grad)

    # endregion static methods
    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.FILL

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
    def prop_angle(self) -> Intensity:
        """Gets/Sets angle of the gradient."""
        pv = cast(int, self._get("Angle"))
        return Intensity(pv)

    @prop_angle.setter
    def prop_angle(self, value: Intensity | int):
        if not isinstance(value, Intensity):
            value = Intensity(value)
        self._set("Angle", value.value * 10)

    @property
    def prop_border(self) -> Intensity:
        """Gets/Sets percent of the total width where just the start color is used."""
        pv = cast(int, self._get("Border"))
        if pv == 0:
            Intensity(0)
        return Intensity(round(pv / 10))

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
