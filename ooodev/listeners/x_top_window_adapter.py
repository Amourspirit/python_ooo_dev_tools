from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.awt import XTopWindowListener
from ooodev.mock import mock_g

if mock_g.DOCS_BUILDING:
    from ooodev.mock import unohelper
else:
    import unohelper


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class XTopWindowAdapter(unohelper.Base, XTopWindowListener):  # type: ignore
    """
    makes it possible to receive window events.

    This class is meant a parent class.
    """

    def windowOpened(self, event: EventObject) -> None:
        """is invoked when a window is activated."""
        pass

    def windowActivated(self, event: EventObject) -> None:
        """is invoked when a window is activated."""
        pass

    def windowDeactivated(self, event: EventObject) -> None:
        """is invoked when a window is deactivated."""
        pass

    def windowMinimized(self, event: EventObject) -> None:
        """is invoked when a window is iconified."""
        pass

    def windowNormalized(self, event: EventObject) -> None:
        """is invoked when a window is deiconified."""
        pass

    def windowClosing(self, event: EventObject) -> None:
        """
        is invoked when a window is in the process of being closed.

        The close operation can be overridden at this point.
        """
        pass

    def windowClosed(self, event: EventObject) -> None:
        """is invoked when a window has been closed."""
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
