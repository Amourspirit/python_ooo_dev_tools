from __future__ import annotations
from typing import TYPE_CHECKING
from ..mock import mock_g

if mock_g.DOCS_BUILDING:
    from ..mock import unohelper
else:
    import unohelper

from com.sun.star.frame import XTerminateListener

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class XTerminateAdapter(unohelper.Base, XTerminateListener):
    """
    XTerminateListener implementation

    See Also:
        `API XTerminateListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XTerminateListener.html>`_
    """

    def notifyTermination(self, Event: EventObject) -> None:
        """
        is called when the master environment is finally terminated.

        No veto will be accepted then.
        """
        pass

    def queryTermination(self, Event: EventObject) -> None:
        """
        is called when the master environment (e.g., desktop) is about to terminate.

        Termination can be intercepted by throwing TerminationVetoException. Interceptor will be the new owner of desktop and should call XDesktop.terminate() after finishing his own operations.

        Raises:
            TerminationVetoException: ``TerminationVetoException``
        """
        pass

    def disposing(self, event: EventObject) -> None:
        """
        gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including XComponent.removeEventListener() ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at XComponent.
        """
        # from com.sun.star.lang.XEventListener
        pass
