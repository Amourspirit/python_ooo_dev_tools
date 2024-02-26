from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.listener_event_args import ListenerEventArgs

from ooodev.adapter.awt.key_handler_events import KeyHandlerEvents
from ooodev.adapter.awt.mouse_click_events import MouseClickEvents
from ooodev.adapter.view.selection_change_events import SelectionChangeEvents
from ooodev.adapter.frame.controller_partial import ControllerPartial
from ooodev.adapter.frame.dispatch_provider_partial import DispatchProviderPartial


if TYPE_CHECKING:
    from com.sun.star.frame import Controller  # service


class ControllerComp(
    ComponentBase,
    ControllerPartial,
    DispatchProviderPartial,
    KeyHandlerEvents,
    MouseClickEvents,
    SelectionChangeEvents,
):
    """
    Class for managing Controller Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Controller) -> None:
        """
        Constructor

        Args:
            component (Controller): UNO Component that implements ``com.sun.star.frame.Controller`` service.
        """
        ComponentBase.__init__(self, component)
        ControllerPartial.__init__(self, component=component, interface=None)
        DispatchProviderPartial.__init__(self, component=component, interface=None)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        KeyHandlerEvents.__init__(self, trigger_args=generic_args, cb=self._on_key_handler_events_add_remove)
        MouseClickEvents.__init__(self, trigger_args=generic_args, cb=self._on_mouse_events_add_remove)
        SelectionChangeEvents.__init__(self, trigger_args=generic_args, cb=self._on_selection_change_events_add_remove)

    # region Lazy Listeners
    def _on_key_handler_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addKeyHandler(self.events_listener_key_handler)
        event.remove_callback = True

    def _on_mouse_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addMouseClickHandler(self.events_listener_mouse_click)
        event.remove_callback = True

    def _on_selection_change_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addSelectionChangeListener(self.events_listener_selection_change)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.frame.Controller",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Controller:
        """Controller Component"""
        # pylint: disable=no-member
        return cast("Controller", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
