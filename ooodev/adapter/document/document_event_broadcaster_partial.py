from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.document import XDocumentEventBroadcaster

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.document import XDocumentEventListener
    from com.sun.star.frame import XController2


class DocumentEventBroadcasterPartial:
    """
    Partial class for XDocumentEventBroadcaster.
    """

    # pylint: disable=unused-argument

    def __init__(
        self, component: XDocumentEventBroadcaster, interface: UnoInterface | None = XDocumentEventBroadcaster
    ) -> None:
        """
        Constructor

        Args:
            component (XDocumentEventBroadcaster): UNO Component that implements ``com.sun.star.container.XDocumentEventBroadcaster`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDocumentEventBroadcaster``.
        """

        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XDocumentEventBroadcaster
    def add_document_event_listener(self, listener: XDocumentEventListener) -> None:
        """
        Registers a listener which is notified about document events
        """
        self.__component.addDocumentEventListener(listener)

    def notify_document_event(self, event_name: str, view_controller: XController2, supplement: Any) -> None:
        """
        Causes the broadcaster to notify all registered listeners of the given event

        The method will create a DocumentEvent instance with the given parameters, and fill in the Source member (denoting the broadcaster) as appropriate.

        Whether the actual notification happens synchronously or asynchronously is up to the implementor of this method. However, implementations are encouraged to specify this, for the list of supported event types, in their service contract.

        Implementations might also decide to limit the list of allowed events (means event names) at their own discretion. Again, in this case they're encouraged to document this in their service contract.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.NoSupportException: ``NoSupportException``
        """
        self.__component.notifyDocumentEvent(event_name, view_controller, supplement)

    def remove_document_event_listener(self, listener: XDocumentEventListener) -> None:
        """
        revokes a listener which has previously been registered to be notified about document events.
        """
        self.__component.removeDocumentEventListener(listener)

    # endregion XDocumentEventBroadcaster
