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
    from ooodev.loader.inst.lo_inst import LoInst


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
        self._on_document_event_add_remove_called = False
        self._enter_count = 0
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
        self._on_document_event_add_remove_called = True
        event.remove_callback = True

    # endregion Lazy Listeners

    # region context manage
    # Context manager temporarily removes the document event listener
    def __enter__(self):
        self._enter_count += 1
        if self._enter_count > 1:
            return self
        if not self._on_document_event_add_remove_called:
            return self
        self.component.removeDocumentEventListener(self.events_listener_document_event)
        return self

    def __exit__(self, *exc) -> None:
        if self._enter_count == 1 and self._on_document_event_add_remove_called:
            self.component.addDocumentEventListener(self.events_listener_document_event)
        self._enter_count -= 1

    # endregion context manage

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

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> TheGlobalEventBroadcasterComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            TheGlobalEventBroadcasterComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        factory = lo_inst.get_singleton("/singletons/com.sun.star.frame.theGlobalEventBroadcaster")  # type: ignore
        if factory is None:
            raise ValueError("Could not get theGlobalEventBroadcaster singleton.")
        return cls(factory)

    # region Properties
    @property
    def component(self) -> theGlobalEventBroadcaster:
        """theGlobalEventBroadcaster Component"""
        # pylint: disable=no-member
        return cast("theGlobalEventBroadcaster", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
