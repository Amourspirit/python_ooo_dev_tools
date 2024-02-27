from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.sheet import XResultListener

from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.generic_args import GenericArgs

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.sheet import ResultEvent
    from com.sun.star.sheet import XVolatileResult


class ResultListener(AdapterBase, XResultListener):
    """
    Allows notification when a new volatile function result is available.

    See Also:
        `API XResultListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XResultListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XVolatileResult | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XVolatileResult, optional): An UNO object that implements the ``XVolatileResult`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addResultListener(self)

    def modified(self, event: ResultEvent) -> None:
        """
        Event is invoked when a new value is available.
        """
        self._trigger_event("modified", event)

    def disposing(self, event: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        self._trigger_event("disposing", event)
