from __future__ import annotations
from typing import Any, cast, Tuple, overload
import uno
from ooo.dyn.awt.point import Point as UnoPoint

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.utils import lo as mLo
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_base import StyleBase
from ooodev.units import UnitT, UnitConvert, UnitMM


class Position(StyleBase):
    """
    Positions a shape.

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        pos_x: float | UnitT,
        pos_y: float | UnitT,
    ) -> None:
        """
        Constructor

        Args:
            pos_x (float, UnitT): Specifies the x-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            pos_y (float, UnitT): Specifies the y-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
        """
        super().__init__()
        # self._chart_doc = chart_doc
        try:
            self._pos_x = pos_x.get_value_mm100()  # type: ignore
        except AttributeError:
            self._pos_x = UnitConvert.convert_mm_mm100(pos_x)  # type: ignore
        try:
            self._pos_y = pos_y.get_value_mm100()  # type: ignore
        except AttributeError:
            self._pos_y = UnitConvert.convert_mm_mm100(pos_y)  # type: ignore

    def _get_property_name(self) -> str:
        return "Position"

    # region Overridden Methods
    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies position properties to ``obj``

        Args:
            obj (Any): UNO object.

        Returns:
            None:
        """
        props = kwargs.pop("override_dv", {})
        update_dv = bool(kwargs.pop("update_dv", True))
        if update_dv:
            name = self._get_property_name()
            if not name:
                return
            struct = UnoPoint(X=self._pos_x, Y=self._pos_y)
            props.update({name: struct})
        # props[name] = struct
        super().apply(obj=obj, override_dv=props)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.Shape",)
        return self._supported_services_values

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
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
        # pylint: disable=protected-access
        cp = cast(Position, super().copy(**kwargs))
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
    def prop_pos_x(self, value: float | UnitT) -> None:
        try:
            self._pos_x = value.get_value_mm100()  # type: ignore
        except AttributeError:
            self._pos_x = UnitConvert.convert_mm_mm100(value)  # type: ignore

    @property
    def prop_pos_y(self) -> UnitMM:
        """Gets or sets the y-coordinate of the position of the shape (in ``mm`` units)."""
        return UnitMM.from_mm100(self._pos_y)

    @prop_pos_y.setter
    def prop_pos_y(self, value: float | UnitT) -> None:
        try:
            self._pos_y = value.get_value_mm100()  # type: ignore
        except AttributeError:
            self._pos_y = UnitConvert.convert_mm_mm100(value)  # type: ignore

    # endregion Properties
