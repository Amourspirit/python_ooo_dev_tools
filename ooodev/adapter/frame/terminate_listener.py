from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from ...events.args.event_args import EventArgs as EventArgs
from ..adapter_base import AdapterBase, GenericArgs as GenericArgs
from ...utils import lo as mLo

from com.sun.star.frame import XTerminateListener

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class TerminateListener(AdapterBase, XTerminateListener):
    """
    Has to be provided if an object wants to receive an event when the master environment (e.g., desktop) is terminated.

    See Also:
        `API XTerminateListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XTerminateListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, add_listener: bool = True) -> None:
        """
        Constructor

        Args:
            add_listener (bool, optional): If ``True`` listener is automatically added. Default ``True``.
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

        if add_listener:
            desktop = mLo.Lo.get_desktop()
            desktop.addTerminateListener(self)

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
