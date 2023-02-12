"""
Module for ``Hatch`` struct.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, Type, cast, overload, TypeVar, TYPE_CHECKING

import uno
from ....exceptions import ex as mEx
from ....utils import props as mProps
from ....utils.color import Color
from ....utils.data_type.angle import Angle as Angle
from ....utils.data_type.intensity import Intensity as Intensity
from ...style_base import StyleBase
from ...kind.format_kind import FormatKind
from ....utils.unit_convert import UnitConvert

from ooo.dyn.drawing.hatch import Hatch
from ooo.dyn.drawing.hatch_style import HatchStyle

# see Also:
# https://github.com/LibreOffice/core/blob/f725629a6241ec064770c28957f11d306c18f130/filter/source/msfilter/escherex.cxx

_THatchStruct = TypeVar(name="_THatchStruct", bound="HatchStruct")


class HatchStruct(StyleBase):
    """
    Represents UNO ``Hatch`` struct.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: HatchStyle = HatchStyle.SINGLE,
        color: Color = Color(0),
        distance: float = 0.0,
        angle: Angle | int = 0,
    ) -> None:
        """
        Constructor

        Args:
            style (HatchStyle, optional): Specifies the kind of lines used to draw this hatch. Default ``HatchStyle.SINGLE``.
            color (Color, optional): Specifies the color of the hatch lines. Default ``0``.
            distance (int, optional): Specifies the distance between the lines in the hatch (in ``mm`` units). Default ``0.0``
            angle (Angle, int, optional): Specifies angle of the hatch in degrees. Default to ``0``.

        Returns:
            None:
        """

        super().__init__()
        self.prop_style = style
        self.prop_color = color
        self.prop_distance = distance
        self.prop_angle = angle

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        return ()

    def _container_get_service_name(self) -> str:
        # https://github.com/LibreOffice/core/blob/d9e044f04ac11b76b9a3dac575f4e9155b67490e/chart2/source/tools/PropertyHelper.cxx#L229
        return "com.sun.star.drawing.HatchTable"

    def _get_property_name(self) -> str:
        return "FillHatch"

    def copy(self: _THatchStruct) -> _THatchStruct:
        nu = super(HatchStruct, self.__class__).__new__(self.__class__)
        nu.__init__()
        if self._dv:
            nu._update(self._dv)
        return nu

    def get_attrs(self) -> Tuple[str, ...]:
        return (self._get_property_name(),)

    def get_uno_struct(self) -> Hatch:
        """
        Gets UNO ``Hatch`` from instance.

        Returns:
            Hatch: ``Hatch`` instance
        """
        return Hatch(
            Style=self._get("Style"),
            Color=self._get("Color"),
            Distance=self._get("Distance"),
            Angle=self._get("Angle"),
        )

    def __eq__(self, oth: object) -> bool:
        obj2 = None
        if isinstance(oth, HatchStruct):
            obj2 = oth.get_uno_struct()
        if getattr(oth, "typeName", None) == "com.sun.star.drawing.Hatch":
            obj2 = cast(Hatch, oth)
        if obj2:
            obj1 = self.get_uno_struct()
            return (
                obj1.Style == obj2.Style
                and obj1.Color == obj2.Color
                and obj1.Distance == obj2.Distance
                and obj1.Angle == obj2.Angle
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
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        if not mProps.Props.has(obj, self._get_property_name()):
            self._print_not_valid_obj("apply")
            return

        hatch = self.get_uno_struct()
        props = {self._get_property_name(): hatch}
        super().apply(obj=obj, override_dv=props)

    # endregion apply()

    # region static methods
    @classmethod
    def from_hatch(cls: Type[_THatchStruct], value: Hatch) -> _THatchStruct:
        """
        Converts a ``Hatch`` instance to a ``HatchStruct``

        Args:
            value (Hatch): UNO Hatch

        Returns:
            HatchStruct: ``HatchStruct`` set with ``Hatch`` properties
        """
        inst = super(HatchStruct, cls).__new__(cls)
        inst.__init__()
        inst._set("Style", value.Style)
        inst._set("Color", value.Color)
        inst._set("Distance", value.Distance)
        inst._set("Angle", value.Angle)
        return inst

    @classmethod
    def from_obj(cls: Type[_THatchStruct], obj: object) -> _THatchStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property

        Returns:
            HatchStruct: ``HatchStruct`` instance that represents ``obj`` hatch properties.
        """
        # this nu is only used to get Property Name
        nu = super(HatchStruct, cls).__new__(cls)
        nu.__init__()
        prop_name = nu._get_property_name()

        try:
            grad = cast(Hatch, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")

        return cls.from_hatch(grad)

    # endregion static methods
    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.TXT_CONTENT

    @property
    def prop_style(self) -> HatchStyle:
        """Gets/Sets the style of the hatch."""
        return self._get("Style")

    @prop_style.setter
    def prop_style(self, value: HatchStyle):
        self._set("Style", value)

    @property
    def prop_color(self) -> Color:
        """Gets/Sets the color of the hatch lines."""
        return self._get("Color")

    @prop_color.setter
    def prop_color(self, value: Color):
        self._set("Color", value)

    @property
    def prop_distance(self) -> float:
        """Gets/Sets the distance between the lines in the hatch (in ``mm`` units)."""
        pv = cast(int, self._get("Distance"))
        return round(UnitConvert.convert_mm100_mm(pv), 2)

    @prop_distance.setter
    def prop_distance(self, value: float):
        self._set("Distance", UnitConvert.convert_mm_mm100(value))

    @property
    def prop_angle(self) -> Angle:
        """Gets/Sets angle of the hatch."""
        pv = cast(int, self._get("Angle"))
        if pv == 0:
            return Angle(0)
        return Angle(round(pv / 10))

    @prop_angle.setter
    def prop_angle(self, value: Angle | int):
        if not isinstance(value, Angle):
            value = Angle(value)
        self._set("Angle", value.value * 10)

    # endregion properties
