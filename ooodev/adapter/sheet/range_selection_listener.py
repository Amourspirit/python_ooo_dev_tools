from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.sheet import XRangeSelectionListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.sheet import RangeSelectionEvent
    from com.sun.star.sheet import XRangeSelection


class RangeSelectionListener(AdapterBase, XRangeSelectionListener):
    """
    Makes it possible to receive events when the active spreadsheet changes.

    **since**

        OOo 2.0

    See Also:
        `API XRangeSelectionListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XActivationEventListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XRangeSelection | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XRangeSelection, optional): An UNO object that implements the ``XRangeSelection`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addRangeSelectionListener(self)

    def aborted(self, event: RangeSelectionEvent) -> None:
        """
        Event is invoked when range selection is aborted.
        """
        self._trigger_event("aborted", event)

    def done(self, event: RangeSelectionEvent) -> None:
        """
        Event is invoked when range selection is completed.
        """
        self._trigger_event("done", event)

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
