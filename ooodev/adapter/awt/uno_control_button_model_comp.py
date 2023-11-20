from __future__ import annotations
from typing import Any, cast, Dict, TYPE_CHECKING
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.lang.event_events import EventEvents
from ooodev.adapter.beans.property_change_collection import PropertyChangeCollection
from ooodev.adapter.beans.vetoable_change_collection import VetoableChangeCollection
from ooodev.adapter.adapter_base import GenericArgs


if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlButtonModel  # service


class UnoControlButtonModelComp(ComponentBase, EventEvents, PropertyChangeCollection, VetoableChangeCollection):
    """
    Class for managing UNO Control Button Model Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: UnoControlButtonModel, generic_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            component (UnoControlButtonModel): UNO Control Button Model Component.
            generic_args (GenericArgs, optional): Generic Arguments to pass to subclass events. Defaults to None.
        """
        ComponentBase.__init__(self, component)
        local_generic_args = self._get_generic_args()
        if generic_args:
            local_generic_args.kwargs.update(generic_args.kwargs)
        EventEvents.__init__(self, trigger_args=local_generic_args, cb=self._on_event_events_add_remove)
        PropertyChangeCollection.__init__(self, component=self.component, generic_args=local_generic_args)
        VetoableChangeCollection.__init__(self, component=self.component, generic_args=local_generic_args)

    # region Lazy Listeners
    def _on_event_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addEventListener(self.events_listener_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlButtonModel",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> UnoControlButtonModel:
        """Tree Data Model Component"""
        return cast("UnoControlButtonModel", self._get_component())

    # endregion Properties
