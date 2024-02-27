from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.util import XRefreshListener

from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.generic_args import GenericArgs

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.util import XRefreshable


class RefreshListener(AdapterBase, XRefreshListener):
    """
    Makes it possible to receive refreshed events.

    See Also:
        `API XRefreshListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XRefreshListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XRefreshable | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XRefreshable, optional): An UNO object that implements the ``XRefreshable`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addRefreshListener(self)

    def refreshed(self, event: EventObject) -> None:
        """
        Event is invoked when when the object data is refreshed.
        """
        self._trigger_event("refreshed", event)

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
