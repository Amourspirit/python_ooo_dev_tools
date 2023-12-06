from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.awt.enhanced_mouse_click_events import EnhancedMouseClickEvents
from ooodev.adapter.awt.key_events import KeyEvents
from ooodev.adapter.awt.mouse_click_events import MouseClickEvents
from ooodev.adapter.view.selection_change_events import SelectionChangeEvents
from .activation_event_events import ActivationEventEvents
from .range_selection_change_events import RangeSelectionChangeEvents
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement

if TYPE_CHECKING:
    from com.sun.star.sheet import SpreadsheetView  # service


class SpreadsheetViewComp(
    ComponentBase,
    ActivationEventEvents,
    EnhancedMouseClickEvents,
    KeyEvents,
    MouseClickEvents,
    RangeSelectionChangeEvents,
    SelectionChangeEvents,
    PropertyChangeImplement,
    VetoableChangeImplement,
):
    """
    Class for managing Spreadsheet View Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: SpreadsheetView) -> None:
        """
        Constructor

        Args:
            component (SpreadsheetView): UNO Spreadsheet View Component
        """
        ComponentBase.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ActivationEventEvents.__init__(self, trigger_args=generic_args, cb=self._on_activation_events_add_remove)
        EnhancedMouseClickEvents.__init__(self, trigger_args=generic_args, cb=self._on_enhanced_mouse_click_add_remove)
        KeyEvents.__init__(self, trigger_args=generic_args, cb=self._on_key_add_remove_add_remove)
        MouseClickEvents.__init__(self, trigger_args=generic_args, cb=self._on_mouse_click_add_remove)
        RangeSelectionChangeEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_range_selection_change_add_remove
        )
        SelectionChangeEvents.__init__(self, trigger_args=generic_args, cb=self._on_selection_change_add_remove)
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Lazy Listeners
    def _on_activation_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addActivationEventListener(self.events_listener_activation_event)
        event.remove_callback = True

    def _on_enhanced_mouse_click_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addEnhancedMouseClickHandler(self.events_listener_enhanced_mouse_click)
        event.remove_callback = True

    def _on_key_add_remove_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addKeyHandler(self.events_listener_key)
        event.remove_callback = True

    def _on_mouse_click_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addMouseClickHandler(self.events_listener_mouse_click)
        event.remove_callback = True

    def _on_range_selection_change_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addRangeSelectionChangeListener(self.events_listener_range_selection_change)
        event.remove_callback = True

    def _on_selection_change_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addSelectionChangeListener(self.events_listener_selection_change)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SpreadsheetView",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> SpreadsheetView:
        """Spreadsheet View Component"""
        return cast("SpreadsheetView", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
