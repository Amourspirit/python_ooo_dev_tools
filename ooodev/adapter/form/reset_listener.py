from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.adapter.adapter_base import AdapterBase, GenericArgs as GenericArgs

from com.sun.star.form import XResetListener

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.form import XReset


class ResetListener(AdapterBase, XResetListener):
    """
    Is the interface for receiving notifications about reset events.

    The listener is called if a component implementing the XReset interface performs a reset.Order of events:

    See Also:
        `API XResetListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1form_1_1XResetListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XReset | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XReset, optional): An UNO object that implements the ``XReset`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)

        if subscriber:
            subscriber.addResetListener(self)

    def approveReset(self, event: EventObject) -> None:
        """
        Event is invoked is invoked before a component is reset.

        No veto will be accepted then.
        """
        self._trigger_event("approveReset", event)

    def resetted(self, event: EventObject) -> None:
        """
        Event is invoked when a component has been reset.
        """
        self._trigger_event("approveReset", event)

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
