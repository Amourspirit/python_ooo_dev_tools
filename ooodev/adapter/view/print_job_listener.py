from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XModel
from com.sun.star.view import XPrintJobListener
from com.sun.star.view import XSelectionSupplier

from ooodev.utils import lo as mLo

from ..adapter_base import AdapterBase, GenericArgs as GenericArgs

if TYPE_CHECKING:
    from com.sun.star.view import PrintJobEvent
    from com.sun.star.view import XPrintJobBroadcaster
    from com.sun.star.lang import EventObject


class PrintJobListener(AdapterBase, XPrintJobListener):
    """
    Receives events about print job progress.

    XPrintJobListener can be registered to ``XPrintJobBroadcaster``.
    Then, the client object will be notified when a new print job starts or its state changes.

    **since**

        OOo 1.1.2

    See Also:
        `API XPrintJobListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1view_1_1XPrintJobListener.html>`_
    """

    def __init__(
        self, trigger_args: GenericArgs | None = None, subscriber: XPrintJobBroadcaster | None = None
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XPrintJobBroadcaster, optional): An UNO object that implements the ``XPrintJobBroadcaster`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addPrintJobListener(self)

    def printJobEvent(self, event: PrintJobEvent) -> None:
        """
        Informs the user about the creation or the progress of a PrintJob.
        """
        self._trigger_event("printJobEvent", event)

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
