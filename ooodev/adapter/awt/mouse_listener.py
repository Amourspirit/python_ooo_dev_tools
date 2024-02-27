from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.awt import XMouseListener
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import MouseEvent
    from com.sun.star.presentation import XSlideShowView
    from com.sun.star.awt import XWindow


class MouseListener(AdapterBase, XMouseListener):
    """
    Makes it possible to receive events from the mouse in a certain window.

    Use the following interfaces which allow to receive (and consume) mouse events even on windows which are not at the top:

    See Also:
        `API XMouseListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMouseListener.html>`_
    """

    def __init__(
        self, trigger_args: GenericArgs | None = None, subscriber: XSlideShowView | XWindow | None = None
    ) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XSlideShowView, XWindow, optional): An UNO object that implements ``XSlideShowView`` or ``XWindow`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addMouseListener(self)

    # region XMouseListener
    def mouseEntered(self, event: MouseEvent) -> None:
        """
        is invoked when the mouse enters a window.
        """
        self._trigger_event("mouseEntered", event)

    def mouseExited(self, event: MouseEvent) -> None:
        """
        is invoked when the mouse exits a window.
        """
        self._trigger_event("mouseExited", event)

    def mousePressed(self, event: MouseEvent) -> None:
        """
        is invoked when a mouse button has been pressed on a window.

        Since mouse presses are usually also used to indicate requests for pop-up menus (also known as context menus) on objects,
        you might receive two events for a single mouse press: For example, if, on your operating system, pressing the right
        mouse button indicates the request for a context menu, then you will receive one call to mousePressed() indicating
        the mouse click, and another one indicating the context menu request. For the latter, the MouseEvent.
        PopupTrigger member of the event will be set to TRUE.
        """
        self._trigger_event("mousePressed", event)

    def mouseReleased(self, event: MouseEvent) -> None:
        """
        is invoked when a mouse button has been released on a window.
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

    # endregion XMouseListener
