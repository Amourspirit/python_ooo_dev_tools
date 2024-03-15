from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.awt.size import Size
from ooodev.adapter.struct_base import StructBase

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT


class SizeStructComp(StructBase[Size]):
    """
    Size Struct.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_awt_Size_changing``.
    The event raised after the property is changed is called ``com_sun_star_awt_Size_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: Size, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (Size): Size.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_awt_Size_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_awt_Size_changed"

    def _copy(self, src: Size | None = None) -> Size:
        if src is None:
            src = self.component
        return Size(
            Width=src.Width,
            Height=src.Height,
        )

    # endregion Overrides

    # region Properties

    @property
    def width(self) -> int:
        """
        Gets/Sets the Width.
        """
        return self.component.Width

    @width.setter
    def width(self, value: int) -> None:
        old_value = self.component.Width
        if old_value != value:
            event_args = self._trigger_cancel_event("Width", old_value, value)
            self._trigger_done_event(event_args)

    @property
    def height(self) -> int:
        """
        Gets/Sets the the Height.
        """
        return self.component.Height

    @height.setter
    def height(self, value: int) -> None:
        old_value = self.component.Height
        if old_value != value:
            event_args = self._trigger_cancel_event("Height", old_value, value)
            self._trigger_done_event(event_args)

    # endregion Properties
