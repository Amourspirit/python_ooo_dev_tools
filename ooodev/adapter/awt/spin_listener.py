from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.awt import XSpinListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import SpinEvent
    from com.sun.star.awt import XSpinField


class SpinListener(AdapterBase, XSpinListener):
    """
    Makes it possible to receive spin events.

    See Also:
        `API XSpinListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XSpinListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XSpinField | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XSpinField, optional): An UNO object that implements the ``XSpinField`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addSpinListener(self)

    # region XSpinListener
    def down(self, event: SpinEvent) -> None:
        """
        Event is invoked when the spin field is spun down.
        """
        self._trigger_event("down", event)

    def first(self, event: SpinEvent) -> None:
        """
        Event is invoked when the spin field is set to the lower value.
        """
        self._trigger_event("first", event)

    def last(self, event: SpinEvent) -> None:
        """
        Event is invoked when the spin field is set to the upper value.
        """
        self._trigger_event("last", event)

    def up(self, event: SpinEvent) -> None:
        """
        Event is invoked when the spin field is spun up.
        """
        self._trigger_event("up", event)

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

    # endregion XSpinListener
