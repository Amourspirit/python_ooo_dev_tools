from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from ...events.args.event_args import EventArgs as EventArgs
from ..adapter_base import AdapterBase, GenericArgs as GenericArgs

from com.sun.star.awt import XWindowListener

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import WindowEvent


class WindowListener(AdapterBase, XWindowListener):
    """
    Makes it possible to receive window events.

    Component events are provided only for notification purposes.
    Moves and resizes will be handled internally by the window component, so that GUI
    layout works properly regardless of whether a program registers such a listener or not.

    See Also:
        `API XWindowListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XWindowListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        super().__init__(trigger_args=trigger_args)

    # region XPaintListener
    def windowHidden(self, event: EventObject) -> None:
        """
        is invoked when the window has been hidden.
        """
        self._trigger_event("windowHidden", event)

    def windowMoved(self, event: WindowEvent) -> None:
        """
        is invoked when the window has been moved.
        """
        self._trigger_event("windowMoved", event)

    def windowResized(self, event: WindowEvent) -> None:
        """
        is invoked when the window has been resized.
        """
        self._trigger_event("windowResized", event)

    def windowShown(self, event: WindowEvent) -> None:
        """
        is invoked when the window has been shown.
        """
        self._trigger_event("windowShown", event)

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

    # endregion XPaintListener
