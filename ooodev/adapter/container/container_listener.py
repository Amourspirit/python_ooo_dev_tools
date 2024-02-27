from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.container import XContainerListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.container import ContainerEvent
    from com.sun.star.container import XContainer


class ContainerListener(AdapterBase, XContainerListener):
    """
    receives events when the content of the related container changes.

    See Also:
        `API XContainerListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1container_1_1XContainerListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XContainer | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XContainer, optional): An UNO object that implements the ``XContainer`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addContainerListener(self)

    # region XContainerListener
    def elementInserted(self, event: ContainerEvent) -> None:
        """
        Event is invoked when a container has inserted an element.
        """
        self._trigger_event("elementInserted", event)

    def elementRemoved(self, event: ContainerEvent) -> None:
        """
        Event is invoked when a container has removed an element.
        """
        self._trigger_event("elementRemoved", event)

    def elementReplaced(self, event: ContainerEvent) -> None:
        """
        Event is invoked when a container has replaced an element.
        """
        self._trigger_event("elementReplaced", event)

    def disposing(self, event: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", event)

    # endregion XContainerListener
