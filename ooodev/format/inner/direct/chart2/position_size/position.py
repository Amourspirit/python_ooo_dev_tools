from __future__ import annotations
from typing import Any, cast, Tuple, overload, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooo.dyn.awt.point import Point as UnoPoint

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.utils import props as mProps
from ooodev.format.inner.style_base import StyleBase
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
    from typing_extensions import Self
else:
    Self = Any


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
    def _container_get_service_name(self) -> str:
        # override to keep type checker happy.
        raise NotImplementedError

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

    # region copy()
    @overload
    def copy(self) -> Position: ...

    @overload
    def copy(self, **kwargs) -> Position: ...

    def copy(self, **kwargs) -> Position:
        """
        Copy the current instance.

        Returns:
            Position: The copied instance.
        """
        # pylint: disable=protected-access
        cp = cast(Position, super().copy(pos_x=0, pos_y=0, **kwargs))
        cp._pos_x = self._pos_x
        cp._pos_y = self._pos_y
        return cp

    # endregion copy()
    # endregion Overridden Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> Self:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Position: New instance.
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> Self:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.
            **kwargs: Additional arguments.

        Returns:
            Position: New instance.
        """
        ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> Self:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Position: New instance.
        """
        # pylint: disable=protected-access
        inst = cls(pos_x=0, pos_y=0, **kwargs)
        name = inst._get_property_name()
        if not name:
            raise ValueError("No property name to retrieve.")

        pt = cast(UnoPoint, mProps.Props.get(obj, name))
        nu = cls(pos_x=UnitMM100(pt.X), pos_y=UnitMM100(pt.Y), **kwargs)
        nu.set_update_obj(obj)
        return nu

    # endregion from_obj()

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
