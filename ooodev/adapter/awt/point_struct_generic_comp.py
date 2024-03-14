from __future__ import annotations
from typing import Generic, TypeVar, TYPE_CHECKING
import uno
from ooo.dyn.awt.point import Point
from ooodev.adapter.struct_base import StructBase
from ooodev.utils.data_type.generic_unit_point import GenericUnitPoint
from ooodev.units.unit_obj import UnitT


if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT

_T = TypeVar("_T", bound="UnitT")

# see ooodev.adapter.drawing.glue_point2_struct_comp.GluePoint2StructComp for example usage.


class PointStructGenericComp(Generic[_T], StructBase[GenericUnitPoint[_T, int]]):
    """
    Generic Point Struct.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``generic_com_sun_star_awt_Point_changing``.
    The event raised after the property is changed is called ``generic_com_sun_star_awt_Point_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(
        self, component: GenericUnitPoint[_T, int], prop_name: str, event_provider: EventsT | None = None
    ) -> None:
        """
        Constructor

        Args:
            component (Point): Point.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "generic_com_sun_star_awt_Point_changing"

    def _get_on_changed_event_name(self) -> str:
        return "generic_com_sun_star_awt_Point_changed"

    def _copy(self, src: Point | None = None) -> Point:
        if src is None:
            src = self.get_uno_point()
        return Point(
            X=src.X,
            Y=src.Y,
        )

    # endregion Overrides

    def get_uno_point(self) -> Point:
        """
        Gets the UNO Point.
        """
        p = self.component.get_point()
        return Point(p.x, p.y)

    # region Properties

    @property
    def x(self) -> _T:
        """
        Gets/Sets the x-coordinate.
        """
        return self.component.x

    @x.setter
    def x(self, value: _T) -> None:
        old_value = self.component.x
        if old_value != value:
            event_args = self._trigger_cancel_event("x", old_value, value)
            self._trigger_done_event(event_args)

    @property
    def y(self) -> _T:
        """
        Gets/Sets the the y-coordinate.
        """
        return self.component.y

    @y.setter
    def y(self, value: int) -> None:
        old_value = self.component.y
        if old_value != value:
            event_args = self._trigger_cancel_event("y", old_value, value)
            self._trigger_done_event(event_args)

    # endregion Properties
