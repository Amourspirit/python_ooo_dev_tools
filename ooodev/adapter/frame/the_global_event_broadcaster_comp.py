from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.listener_event_args import ListenerEventArgs

from ooodev.adapter.document.events_supplier_partial import EventsSupplierPartial
from ooodev.adapter.document.document_event_broadcaster_partial import DocumentEventBroadcasterPartial
from ooodev.adapter.container.set_partial import SetPartial
from ooodev.adapter.document.document_event_events import DocumentEventEvents

if TYPE_CHECKING:
    from com.sun.star.frame import theGlobalEventBroadcaster  # singleton
    from com.sun.star.document import DocumentEvent


class TheGlobalEventBroadcasterComp(
    ComponentBase, EventsSupplierPartial, DocumentEventBroadcasterPartial, SetPartial, DocumentEventEvents
):
    """
    Class for managing theGlobalEventBroadcaster Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: theGlobalEventBroadcaster) -> None:
        """
        Constructor

        Args:
            component (theGlobalEventBroadcaster): UNO Component that implements ``com.sun.star.frame.theGlobalEventBroadcaster`` service.
        """
        ComponentBase.__init__(self, component)
        EventsSupplierPartial.__init__(self, component=component, interface=None)
        DocumentEventBroadcasterPartial.__init__(self, component=component, interface=None)
        SetPartial.__init__(self, component=component, interface=None)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        DocumentEventEvents.__init__(self, trigger_args=generic_args, cb=self._on_document_event_add_remove)

    # region Lazy Listeners
    def _on_document_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addDocumentEventListener(self.events_listener_document_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region XDocumentEventListener
    def document_event_occured(self, event: DocumentEvent) -> None:
        """
        Event is invoked when a document event occurred
        """
        self.component.documentEventOccured(event)

    # endregion XDocumentEventListener

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> theGlobalEventBroadcaster:
        """theGlobalEventBroadcaster Component"""
        # pylint: disable=no-member
        return cast("theGlobalEventBroadcaster", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
