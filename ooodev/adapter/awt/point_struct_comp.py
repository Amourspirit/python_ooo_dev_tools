from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.awt.point import Point
from ooodev.adapter.struct_base import StructBase
from ooodev.utils.data_type.intensity import Intensity

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT


class PointStructComp(StructBase[Point]):
    """
    Point Struct.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_awt_Point_changing``.
    The event raised after the property is changed is called ``com_sun_star_awt_Point_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: Point, prop_name: str, event_provider: EventsT | None = None) -> None:
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
        return "com_sun_star_awt_Point_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_awt_Point_changed"

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
    def x(self) -> int:
        """
        Gets/Sets the x-coordinate.
        """
        return self.component.X

    @x.setter
    def x(self, value: int) -> None:
        old_value = self.component.X
        if old_value != value:
            event_args = self._trigger_cancel_event("X", old_value, value)
            self._trigger_done_event(event_args)

    @property
    def y(self) -> int:
        """
        Gets/Sets the the y-coordinate.
        """
        return self.component.Y

    @y.setter
    def y(self, value: int) -> None:
        val = Intensity(value).value
        old_value = self.component.Y
        if old_value != val:
            event_args = self._trigger_cancel_event("Y", old_value, val)
            self._trigger_done_event(event_args)

    # endregion Properties
