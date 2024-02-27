from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.awt.tab import XTabPageContainerListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt.tab import TabPageActivatedEvent
    from com.sun.star.awt.tab import XTabPageContainer


class TabPageContainerListener(AdapterBase, XTabPageContainerListener):
    """
    An instance of this interface is used by the XTabPageContainer to get notifications about changes in activation of tab pages.

    **since**

        OOo 3.4

    See Also:
        `API XTabPageContainerListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1tab_1_1XTabPageContainerListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XTabPageContainer | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XTabPageContainer, optional): An UNO object that implements the ``XTabPageContainer`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addTabPageContainerListener(self)

    # region XTabPageContainerListener
    def tabPageActivated(self, event: TabPageActivatedEvent) -> None:
        """
        Invoked after a tab page was activated.
        """
        self._trigger_event("tabPageActivated", event)

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

    # endregion XTabPageContainerListener
