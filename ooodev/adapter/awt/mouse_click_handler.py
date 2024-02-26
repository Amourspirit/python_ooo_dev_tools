from __future__ import annotations
import contextlib
from typing import TYPE_CHECKING

import uno
from com.sun.star.awt import XMouseClickHandler

from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.generic_args import GenericArgs

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import MouseEvent
    from com.sun.star.awt import XUserInputInterception


# This class unlike most listeners does not end with the name Listeners


class MouseClickHandler(AdapterBase, XMouseClickHandler):
    """
    Makes it possible to receive events from the mouse in a certain window.

    **since**

        OOo 1.1.2

    See Also:
        `API XMouseClickHandler <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMouseClickHandler.html>`_
    """

    def __init__(
        self, trigger_args: GenericArgs | None = None, subscriber: XUserInputInterception | None = None
    ) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XUserInputInterception, optional): An UNO object that implements the ``XUserInputInterception`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            with contextlib.suppress(AttributeError):
                subscriber.addMouseClickHandler(self)

    # region XMouseClickHandler
    def mousePressed(self, event: MouseEvent) -> None:
        """
        Event is invoked when a mouse button has been pressed on a window.
        """
        self._trigger_event("mousePressed", event)

    def mouseReleased(self, event: MouseEvent) -> None:
        """
        Event is invoked when a mouse button has been released on a window.
        """
        self._trigger_event("mouseReleased", event)

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

    # endregion XMouseClickHandler
