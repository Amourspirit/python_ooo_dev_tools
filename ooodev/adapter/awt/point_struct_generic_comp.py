from __future__ import annotations
from typing import cast, Generic, Type, TypeVar, TYPE_CHECKING
import uno
from ooo.dyn.awt.point import Point
from ooodev.adapter.struct_base import StructBase
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.units.unit_factory import get_unit


if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.units.unit_obj import UnitT

_T = TypeVar("_T", bound="UnitT")

# see ooodev.adapter.drawing.glue_point2_struct_comp.GluePoint2StructComp for example usage.


class PointStructGenericComp(StructBase[Point], Generic[_T]):
    """
    Generic Point Struct.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``generic_com_sun_star_awt_Point_changing``.
    The event raised after the property is changed is called ``generic_com_sun_star_awt_Point_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(
        self, component: Point, unit: Type[_T], prop_name: str, event_provider: EventsT | None = None
    ) -> None:
        """
        Constructor

        Args:
            component (Point): Point.
            unit (Type[UnitT]): Unit Type.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)
        self._unit = unit
        self._unit_length = unit.get_unit_length()
        self._require_convert = self._unit_length != UnitMM100.get_unit_length()

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}[{self._unit.__name__}] {repr(self.component)}"

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "generic_com_sun_star_awt_Point_changing"

    def _get_on_changed_event_name(self) -> str:
        return "generic_com_sun_star_awt_Point_changed"

    def _copy(self, src: Point | None = None) -> Point:
        if src is None:
            src = self.component
        return Point(
            X=src.X,
            Y=src.Y,
        )

    # endregion Overrides

    # region Properties

    @property
    def x(self) -> _T:
        """
        Gets/Sets the x-coordinate.

        When setting the value can be a ``int`` in ``1/100th mm`` units or a ``UnitT`` measurement unit.

        Returns:
            _T: ``UnitT`` measurement unit.
        """
        unit100 = UnitMM100(self.component.X)
        if not self._require_convert:
            return cast(_T, unit100)
        val = unit100.convert_to(self._unit_length)
        return cast(_T, get_unit(self._unit_length, val))

    @x.setter
    def x(self, value: _T | float) -> None:
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        old_value = self.component.X
        if old_value != new_value:
            event_args = self._trigger_cancel_event("X", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def y(self) -> _T:
        """
        Gets/Sets the the y-coordinate.

        When setting the value can be a ``int`` in ``1/100th mm`` units or a ``UnitT`` measurement unit.

        Returns:
            _T: ``UnitT`` measurement unit.
        """
        unit100 = UnitMM100(self.component.Y)
        if not self._require_convert:
            return cast(_T, unit100)
        val = unit100.convert_to(self._unit_length)
        return cast(_T, get_unit(self._unit_length, val))

    @y.setter
    def y(self, value: _T | float) -> None:
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        old_value = self.component.Y
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Y", old_value, new_value)
            self._trigger_done_event(event_args)

    # endregion Properties
