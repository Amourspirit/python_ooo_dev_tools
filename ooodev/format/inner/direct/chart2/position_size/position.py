from __future__ import annotations
from typing import Any, Tuple, overload
import uno

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.utils import lo as mLo
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_base import StyleBase
from ooodev.units import UnitObj, UnitConvert, UnitMM
from ooodev.format.inner.direct.structs.point_struct import PointStruct


class Position(StyleBase):
    """
    Positions a shape.

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        pos_x: float | UnitObj,
        pos_y: float | UnitObj,
    ) -> None:
        """
        Constructor

        Args:
            pos_x (float | UnitObj): Specifies the x-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            pos_y (float | UnitObj): Specifies the y-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
        """
        super().__init__()
        # self._chart_doc = chart_doc
        try:
            self._pos_x = pos_x.get_value_mm100()
        except AttributeError:
            self._pos_x = UnitConvert.convert_mm_mm100(pos_x)
        try:
            self._pos_y = pos_y.get_value_mm100
        except AttributeError:
            self._pos_y = UnitConvert.convert_mm_mm100(pos_y)

    def _get_property_name(self) -> str:
        return "Position"

    # region Overridden Methods
    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies tab properties to ``obj``

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        name = self._get_property_name()
        if not name:
            return
        ps = PointStruct(x=self._pos_x, y=self._pos_y)
        struct = ps.get_uno_struct()
        props = {name: struct}
        super().apply(obj=obj, override_dv=props)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.Shape",)
        return self._supported_services_values

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"PointStruct.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # region copy()
    @overload
    def copy(self) -> Position:
        ...

    @overload
    def copy(self, **kwargs) -> Position:
        ...

    def copy(self, **kwargs) -> Position:
        """
        Copy the current instance.

        Returns:
            Position: The copied instance.
        """
        cp = super().copy(**kwargs)
        cp._pos_x = self._pos_x
        cp._pos_y = self._pos_y
        return cp

    # endregion copy()
    # endregion Overridden Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop

    @property
    def prop_pos_x(self) -> UnitMM:
        """Gets or sets the x-coordinate of the position of the shape (in ``mm`` units)."""
        return UnitMM.from_mm100(self._pos_x)

    @prop_pos_x.setter
    def prop_pos_x(self, value: float | UnitObj) -> None:
        try:
            self._pos_x = value.get_value_mm100()
        except AttributeError:
            self._pos_x = UnitConvert.convert_mm_mm100(value)

    @property
    def prop_pos_y(self) -> UnitMM:
        """Gets or sets the y-coordinate of the position of the shape (in ``mm`` units)."""
        return UnitMM.from_mm100(self._pos_y)

    @prop_pos_y.setter
    def prop_pos_y(self, value: float | UnitObj) -> None:
        try:
            self._pos_y = value.get_value_mm100()
        except AttributeError:
            self._pos_y = UnitConvert.convert_mm_mm100(value)

    # endregion Properties
