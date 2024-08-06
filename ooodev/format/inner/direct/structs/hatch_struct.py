"""
Module for ``Hatch`` struct.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Tuple, Type, cast, overload, TypeVar, TYPE_CHECKING

import uno
from ooo.dyn.drawing.hatch import Hatch
from ooo.dyn.drawing.hatch_style import HatchStyle

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.structs.struct_base import StructBase
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.units.angle import Angle
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_mm import UnitMM
from ooodev.utils import props as mProps
from ooodev.utils.color import Color


if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
# endregion Import

# see Also:
# https://github.com/LibreOffice/core/blob/f725629a6241ec064770c28957f11d306c18f130/filter/source/msfilter/escherex.cxx

_THatchStruct = TypeVar("_THatchStruct", bound="HatchStruct")


class HatchStruct(StructBase):
    """
    Represents UNO ``Hatch`` struct.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: HatchStyle = HatchStyle.SINGLE,
        color: Color = Color(0),
        distance: float | UnitT = 0.0,
        angle: Angle | int = 0,
    ) -> None:
        """
        Constructor

        Args:
            style (HatchStyle, optional): Specifies the kind of lines used to draw this hatch. Default ``HatchStyle.SINGLE``.
            color (:py:data:`~.utils.color.Color`, optional): Specifies the color of the hatch lines. Default ``0``.
            distance (int, UnitT, optional): Specifies the distance between the lines in the hatch (in ``mm`` units) or :ref:`proto_unit_obj`. Default ``0.0``
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
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.Style", "com.sun.star.text.TextFrame")
        return self._supported_services_values

    def _container_get_service_name(self) -> str:
        # https://github.com/LibreOffice/core/blob/d9e044f04ac11b76b9a3dac575f4e9155b67490e/chart2/source/tools/PropertyHelper.cxx#L229
        return "com.sun.star.drawing.HatchTable"

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "FillHatch"
        return self._property_name

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
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies tab properties to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        if not mProps.Props.has(obj, self._get_property_name()):
            self._print_not_valid_srv("apply")
            return

        hatch = self.get_uno_struct()
        props = {self._get_property_name(): hatch}
        super().apply(obj=obj, override_dv=props)

    # endregion apply()

    # region static methods

    # region from_hatch()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_THatchStruct], value: Hatch) -> _THatchStruct: ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_THatchStruct], value: Hatch, **kwargs) -> _THatchStruct: ...

    @classmethod
    def from_uno_struct(cls: Type[_THatchStruct], value: Hatch, **kwargs) -> _THatchStruct:
        """
        Converts a ``Hatch`` instance to a ``HatchStruct``

        Args:
            value (Hatch): UNO Hatch

        Returns:
            HatchStruct: ``HatchStruct`` set with ``Hatch`` properties
        """
        inst = cls(**kwargs)
        inst._set("Style", value.Style)
        inst._set("Color", value.Color)
        inst._set("Distance", value.Distance)
        inst._set("Angle", value.Angle)
        return inst

    # endregion from_hatch()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_THatchStruct], obj: Any) -> _THatchStruct: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_THatchStruct], obj: Any, **kwargs) -> _THatchStruct: ...

    @classmethod
    def from_obj(cls: Type[_THatchStruct], obj: Any, **kwargs) -> _THatchStruct:
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
        nu = cls(**kwargs)
        prop_name = nu._get_property_name()

        try:
            grad = cast(Hatch, mProps.Props.get(obj, prop_name))
        except mEx.PropertyNotFoundError as e:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property") from e

        return cls.from_uno_struct(grad, **kwargs)

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
    def prop_distance(self) -> UnitMM:
        """Gets/Sets the distance between the lines in the hatch (in ``mm`` units)."""
        pv = cast(int, self._get("Distance"))
        return UnitMM(round(UnitConvert.convert_mm100_mm(pv), 2))

    @prop_distance.setter
    def prop_distance(self, value: float | UnitT):
        try:
            self._set("Distance", value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set("Distance", UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_angle(self) -> Angle:
        """Gets/Sets angle of the hatch."""
        pv = cast(int, self._get("Angle"))
        return Angle(0) if pv == 0 else Angle(round(pv / 10))

    @prop_angle.setter
    def prop_angle(self, value: Angle | int):
        if not isinstance(value, Angle):
            value = Angle(value)
        self._set("Angle", value.value * 10)

    # endregion properties
