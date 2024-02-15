from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.frame import XGlobalEventBroadcaster

from ooodev.adapter.document.events_supplier_partial import EventsSupplierPartial
from ooodev.adapter.document.document_event_broadcaster_partial import DocumentEventBroadcasterPartial
from ooodev.adapter.container.set_partial import SetPartial

if TYPE_CHECKING:
    from com.sun.star.document import DocumentEvent
    from ooodev.utils.type_var import UnoInterface


class GlobalEventBroadcasterPartial(EventsSupplierPartial, DocumentEventBroadcasterPartial, SetPartial):
    """
    Partial class for XGlobalEventBroadcaster.
    """

    # pylint: disable=unused-argument

    def __init__(
        self, component: XGlobalEventBroadcaster, interface: UnoInterface | None = XGlobalEventBroadcaster
    ) -> None:
        """
        Constructor

        Args:
            component (XGlobalEventBroadcaster): UNO Component that implements ``com.sun.star.frame.XGlobalEventBroadcaster``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XGlobalEventBroadcaster``.
        """
        EventsSupplierPartial.__init__(self, component, interface)
        DocumentEventBroadcasterPartial.__init__(self, component, interface)
        SetPartial.__init__(self, component, interface)
        self.__component = component

    # region XDocumentEventListener
    def document_event_occured(self, event: DocumentEvent) -> None:
        """
        Event is invoked when a document event occurred
        """
        self.__component.documentEventOccured(event)

    # endregion XDocumentEventListener
