from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.form.binding.list_entry_events import ListEntryEvents
from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.lang.event_events import EventEvents

if TYPE_CHECKING:
    from com.sun.star.table import CellRangeListSource  # service


class CellRangeListSourceComp(ComponentBase, ListEntryEvents, EventEvents):
    """
    Class for managing table CellRangeListSource Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: CellRangeListSource) -> None:
        """
        Constructor

        Args:
            component (CellRangeListSource): UNO table CellRangeListSource Component.
        """
        ComponentBase.__init__(self, component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ListEntryEvents.__init__(self, trigger_args=generic_args, cb=self._on_list_entry_add_remove)
        EventEvents.__init__(self, trigger_args=generic_args, cb=self._on_event_add_remove)

    # region Lazy Listeners
    def _on_list_entry_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addListEntryListener(self.events_listener_list_entry)
        event.remove_callback = True

    def _on_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addEventListener(self.events_listener_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.table.CellRangeListSource",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> CellRangeListSource:
        """CellRangeListSource Component"""
        # pylint: disable=no-member
        return cast("CellRangeListSource", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
