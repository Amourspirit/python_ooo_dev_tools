from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.util.refresh_events import RefreshEvents


if TYPE_CHECKING:
    from com.sun.star.sheet import SheetLink  # service


class SheetLinkComp(ComponentBase, RefreshEvents):
    """
    Class for managing Sheet Link Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: SheetLink) -> None:
        """
        Constructor

        Args:
            component (SheetLink): UNO Sheet Link Component
        """
        ComponentBase.__init__(self, component)
        generic_args = self._get_generic_args()
        RefreshEvents.__init__(self, trigger_args=generic_args, cb=self._on_sheet_link_events_add_remove)

    # region Lazy Listeners
    def _on_sheet_link_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addRefreshListener(self.events_listener_refresh)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SheetLink",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> SheetLink:
        """Tree Data Model Component"""
        return cast("SheetLink", self._get_component())

    # endregion Properties
