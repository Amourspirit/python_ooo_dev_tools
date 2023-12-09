from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.lang.event_events import EventEvents


if TYPE_CHECKING:
    from com.sun.star.table import ListPositionCellBinding  # service


class ListPositionCellBindingComp(
    ComponentBase, ModifyEvents, PropertyChangeImplement, VetoableChangeImplement, EventEvents
):
    """
    Class for managing table ListPositionCellBinding Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ListPositionCellBinding) -> None:
        """
        Constructor

        Args:
            component (ListPositionCellBinding): UNO table ListPositionCellBinding Component.
        """
        ComponentBase.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_add_remove)
        EventEvents.__init__(self, trigger_args=generic_args, cb=self._on_event_add_remove)

    # region Lazy Listeners
    def _on_modify_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addModifyListener(self.events_listener_modify)
        event.remove_callback = True

    def _on_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addEventListener(self.events_listener_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.table.ListPositionCellBinding",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> ListPositionCellBinding:
        """ListPositionCellBinding Component"""
        return cast("ListPositionCellBinding", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
