from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.frame import XTerminateListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.frame import XDesktop


class TerminateListener(AdapterBase, XTerminateListener):
    """
    Has to be provided if an object wants to receive an event when the master environment (e.g., desktop) is terminated.

    See Also:
        - :ref:`ch04_detect_end`
        - :ref:`ch25_listening_close_box`
        - `API XTerminateListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XTerminateListener.html>`_
    """

    def __init__(
        self, trigger_args: GenericArgs | None = None, add_listener: bool = True, subscriber: XDesktop | None = None
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            add_listener (bool, optional): If ``True`` listener is automatically added. Default ``True``.
            subscriber (XDesktop, optional): An UNO object that implements the ``XDesktop`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)

        if add_listener:
            desktop = mLo.Lo.get_desktop()
            desktop.addTerminateListener(self)

        if subscriber:
            subscriber.addTerminateListener(self)

    def notifyTermination(self, event: EventObject) -> None:
        """
        Is called when the master environment is finally terminated.

        No veto will be accepted then.
        """
        self._trigger_event("notifyTermination", event)

    def queryTermination(self, event: EventObject) -> None:
        """
        Is called when the master environment (e.g., desktop) is about to terminate.

        Termination can be intercepted by throwing TerminationVetoException. Interceptor will be the new owner of desktop and should call XDesktop.terminate() after finishing his own operations.

        Raises:
            TerminationVetoException: ``TerminationVetoException``
        """
        self._trigger_event("queryTermination", event)

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
