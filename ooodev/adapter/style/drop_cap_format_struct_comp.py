from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.style.drop_cap_format import DropCapFormat
from ooodev.adapter.struct_base import StructBase
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.units.unit_obj import UnitT


class DropCapFormatStructComp(StructBase[DropCapFormat]):
    """
    Drop Cap Format Struct.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_style_DropCapFormat_changing``.
    The event raised after the property is changed is called ``com_sun_star_style_DropCapFormat_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: DropCapFormat, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (DropCapFormat): Drop Cap Format Component.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_style_DropCapFormat_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_style_DropCapFormat_changed"

    def _copy(self, src: DropCapFormat | None = None) -> DropCapFormat:
        if src is None:
            src = self.component
        return DropCapFormat(
            Lines=src.Lines,
            Count=src.Count,
            Distance=src.Distance,
        )

    # endregion Overrides

    # region Properties

    @property
    def lines(self) -> int:
        """
        This is the number of lines used for a drop cap.
        """
        return self.component.Lines

    @lines.setter
    def lines(self, value: int) -> None:
        old_value = self.component.Lines
        new_value = value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Lines", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def count(self) -> int:
        """
        This is the number of characters in the drop cap.
        """
        return self.component.Count

    @count.setter
    def count(self, value: int) -> None:
        old_value = self.component.Count
        new_value = value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Count", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def distance(self) -> UnitMM100:
        """
        Gets/Sets the distance between the drop cap in the following text.

        Returns:
            UnitMM100: Distance in 100th of a millimeter.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``
        """
        return UnitMM100(self.component.Distance)

    @distance.setter
    def distance(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        old_value = self.component.Distance
        new_value = val.value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Distance", old_value, new_value)
            self._trigger_done_event(event_args)

    # endregion Properties
