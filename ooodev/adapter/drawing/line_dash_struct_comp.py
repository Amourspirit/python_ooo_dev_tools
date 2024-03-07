from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.drawing.line_dash import LineDash
from ooo.dyn.drawing.dash_style import DashStyle

from ooodev.adapter.struct_base import StructBase

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT


class LineDashStructComp(StructBase[LineDash]):
    """
    LineDash Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_drawing_LineDash_changing``.
    The event raised after the property is changed is called ``com_sun_star_drawing_LineDash_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: LineDash, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (LineDash): Border Line.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_drawing_LineDash_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_drawing_LineDash_changed"

    def _copy(self, src: LineDash | None = None) -> LineDash:
        if src is None:
            src = self.component
        return LineDash(
            Style=src.Style,
            Dots=src.Dots,
            DotLen=src.DotLen,
            Dashes=src.Dashes,
            DashLen=src.DashLen,
            Distance=src.Distance,
        )

    # endregion Overrides

    # region Properties
    @property
    def style(self) -> DashStyle:
        """
        Gets/Sets the style of the line dash.

        Returns:
            DashStyle: LineDash Style

        Hint:
            - ``DashStyle`` can be imported from ``ooodev.drawing.dash_style``
        """
        return self.component.Style  # type: ignore

    @style.setter
    def style(self, value: DashStyle) -> None:
        old_value = self.component.Style
        if old_value != value:
            event_args = self._trigger_cancel_event("Style", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def dots(self) -> int:
        """
        Gets/Sets the number of dots in this LineDash.
        """
        return self.component.Dots

    @dots.setter
    def dots(self, value: int) -> None:
        old_value = self.component.Dots
        if old_value != value:
            event_args = self._trigger_cancel_event("Dots", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def dot_len(self) -> int:
        """
        Gets/Sets the length of a dot.
        """
        return self.component.DotLen

    @dot_len.setter
    def dot_len(self, value: int) -> None:
        old_value = self.component.DotLen
        if old_value != value:
            event_args = self._trigger_cancel_event("DotLen", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def dashes(self) -> int:
        """
        Gets/Sets the number of dashes.
        """
        return self.component.Dashes

    @dashes.setter
    def dashes(self, value: int) -> None:
        old_value = self.component.Dashes
        if old_value != value:
            event_args = self._trigger_cancel_event("Dashes", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def dash_len(self) -> int:
        """
        Gets/Sets the length of a single dash.
        """
        return self.component.DashLen

    @dash_len.setter
    def dash_len(self, value: int) -> None:
        old_value = self.component.DashLen
        if old_value != value:
            event_args = self._trigger_cancel_event("DashLen", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def distance(self) -> int:
        """
        Gets/Sets the distance between the dots.
        """
        return self.component.Distance

    @distance.setter
    def distance(self, value: int) -> None:
        old_value = self.component.Distance
        if old_value != value:
            event_args = self._trigger_cancel_event("Distance", old_value, value)
            _ = self._trigger_done_event(event_args)

    # endregion Properties
