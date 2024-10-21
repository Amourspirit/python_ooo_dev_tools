from __future__ import annotations
from typing import TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

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

    @override
    def windowOpened(self, e: EventObject) -> None:
        """is invoked when a window is activated."""
        pass

    @override
    def windowActivated(self, e: EventObject) -> None:
        """is invoked when a window is activated."""
        pass

    @override
    def windowDeactivated(self, e: EventObject) -> None:
        """is invoked when a window is deactivated."""
        pass

    @override
    def windowMinimized(self, e: EventObject) -> None:
        """is invoked when a window is iconified."""
        pass

    @override
    def windowNormalized(self, e: EventObject) -> None:
        """is invoked when a window is deiconified."""
        pass

    @override
    def windowClosing(self, e: EventObject) -> None:
        """
        is invoked when a window is in the process of being closed.

        The close operation can be overridden at this point.
        """
        pass

    @override
    def windowClosed(self, e: EventObject) -> None:
        """is invoked when a window has been closed."""
        pass

    @override
    def disposing(self, Source: EventObject) -> None:
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
