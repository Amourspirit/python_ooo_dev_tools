from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.sheet import XRangeSelectionChangeListener

from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.generic_args import GenericArgs

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.sheet import RangeSelectionEvent
    from com.sun.star.sheet import XRangeSelection


class RangeSelectionChangeListener(AdapterBase, XRangeSelectionChangeListener):
    """
    allows notification when the selected range is changed.

    See Also:
        `API XRangeSelectionChangeListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XRangeSelectionChangeListener.html>`_
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
            subscriber.addRangeSelectionChangeListener(self)

    def descriptorChanged(self, event: RangeSelectionEvent) -> None:
        """
        Event is invoked when the selected range is changed while range selection is active.
        """
        self._trigger_event("descriptorChanged", event)

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
