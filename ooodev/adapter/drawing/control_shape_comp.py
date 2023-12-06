from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.lang.event_events import EventEvents
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement


if TYPE_CHECKING:
    from com.sun.star.drawing import ControlShape  # service


class ControlShapeComp(ComponentBase, EventEvents, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing Volatile Result Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ControlShape) -> None:
        """
        Constructor

        Args:
            component (ControlShape): UNO Volatile Result Component
        """
        ComponentBase.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        EventEvents.__init__(self, trigger_args=generic_args, cb=self._on_event_add_remove)
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Lazy Listeners
    def _on_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addEventListener(self.events_listener_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.ControlShape",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> ControlShape:
        """Volatile Result Component"""
        return cast("ControlShape", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
