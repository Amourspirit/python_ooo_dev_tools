from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.text.table_column_separator import TableColumnSeparator
from ooodev.adapter.struct_base import StructBase

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT

# It seems that it is necessary to assign the struct to a variable, then change the variable and assign it back to the component.
# It is as if LibreOffice creates a new instance of the struct when it is changed.


class TableColumnSeparatorStructComp(StructBase[TableColumnSeparator]):
    """
    Table Column Separator Struct.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_text_TableColumnSeparator_changing``.
    The event raised after the property is changed is called ``com_sun_star_text_TableColumnSeparator_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: TableColumnSeparator, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (TableColumnSeparator): Table Column Separator.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_text_TableColumnSeparator_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_text_TableColumnSeparator_changed"

    def _get_prop_name(self) -> str:
        return self._prop_name

    def _copy(self, src: TableColumnSeparator | None = None) -> TableColumnSeparator:
        if src is None:
            src = self.component
        return TableColumnSeparator(
            Position=src.Position,
            IsVisible=src.IsVisible,
        )

    # endregion Overrides

    # region Properties
    @property
    def position(self) -> int:
        """
        Gets/Sets the position of the separator.
        """
        return self.component.Position

    @position.setter
    def position(self, value: int) -> None:
        old_value = self.component.Position
        if old_value != value:
            event_args = self._trigger_cancel_event("Position", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def is_visible(self) -> bool:
        """
        Gets/Sets if the separator is visible.

        Returns:
            bool: Is Visible.
        """
        return self.component.IsTransparent  # type: ignore

    @is_visible.setter
    def is_visible(self, value: bool) -> None:
        old_value = self.component.IsVisible
        if old_value != value:
            event_args = self._trigger_cancel_event("IsVisible", old_value, value)
            _ = self._trigger_done_event(event_args)

    # endregion Properties
