from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.awt.enhanced_mouse_click_events import EnhancedMouseClickEvents
from ooodev.adapter.awt.key_events import KeyEvents
from ooodev.adapter.awt.mouse_click_events import MouseClickEvents
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.adapter.view.selection_change_events import SelectionChangeEvents
from ooodev.adapter.view.selection_supplier_partial import SelectionSupplierPartial
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.sheet.activation_broadcaster_partial import ActivationBroadcasterPartial
from ooodev.adapter.sheet.activation_event_events import ActivationEventEvents
from ooodev.adapter.sheet.enhanced_mouse_click_broadcaster_partial import EnhancedMouseClickBroadcasterPartial
from ooodev.adapter.sheet.range_selection_change_events import RangeSelectionChangeEvents
from ooodev.adapter.sheet.range_selection_partial import RangeSelectionPartial
from ooodev.adapter.sheet.spreadsheet_view_pane_comp import SpreadsheetViewPaneComp
from ooodev.adapter.sheet.spreadsheet_view_partial import SpreadsheetViewPartial
from ooodev.adapter.sheet.view_freezable_partial import ViewFreezablePartial
from ooodev.adapter.sheet.view_splitable_partial import ViewSplitablePartial

if TYPE_CHECKING:
    from com.sun.star.sheet import SpreadsheetView  # service


class SpreadsheetViewComp(
    SpreadsheetViewPaneComp,
    ActivationBroadcasterPartial,
    ActivationEventEvents,
    EnhancedMouseClickBroadcasterPartial,
    EnumerationAccessPartial,
    IndexAccessPartial,
    RangeSelectionPartial,
    SelectionSupplierPartial,
    SpreadsheetViewPartial,
    ViewFreezablePartial,
    ViewSplitablePartial,
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
        SpreadsheetViewPaneComp.__init__(self, component)
        ActivationBroadcasterPartial.__init__(self, component=component, interface=None)
        EnhancedMouseClickBroadcasterPartial.__init__(self, component=component, interface=None)
        EnumerationAccessPartial.__init__(self, component=component, interface=None)
        IndexAccessPartial.__init__(self, component=component, interface=None)
        RangeSelectionPartial.__init__(self, component=component, interface=None)
        SelectionSupplierPartial.__init__(self, component=component, interface=None)
        SpreadsheetViewPartial.__init__(self, component=component, interface=None)
        ViewFreezablePartial.__init__(self, component=component, interface=None)
        ViewSplitablePartial.__init__(self, component=component, interface=None)
        # pylint: disable=no-member
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
    if TYPE_CHECKING:

        @property
        def component(self) -> SpreadsheetView:
            """Spreadsheet View Component"""
            # pylint: disable=no-member
            return cast("SpreadsheetView", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
