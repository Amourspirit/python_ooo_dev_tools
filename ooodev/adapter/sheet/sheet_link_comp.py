from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.util.refresh_events import RefreshEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs


if TYPE_CHECKING:
    from com.sun.star.sheet import SheetLink  # service


class SheetLinkComp(ComponentBase, RefreshEvents, PropertyChangeImplement, VetoableChangeImplement):
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
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        RefreshEvents.__init__(self, trigger_args=generic_args, cb=self._on_sheet_link_events_add_remove)
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Lazy Listeners
    def _on_sheet_link_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addRefreshListener(self.events_listener_refresh)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SheetLink",)

    # endregion Overrides
    # region Properties
    @property
    @override
    def component(self) -> SheetLink:
        """Sheet Link Component"""
        # pylint: disable=no-member
        return cast("SheetLink", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
