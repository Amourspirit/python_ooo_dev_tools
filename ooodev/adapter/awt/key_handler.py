from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.generic_args import GenericArgs

from com.sun.star.awt import XKeyHandler

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import KeyEvent
    from com.sun.star.awt import XUserInputInterception


class KeyHandler(AdapterBase, XKeyHandler):
    """
    This key handler is similar to ``com.sun.star.awt.XKeyListener`` but allows the consumption of key events.

    If a key event is consumed by one handler both the following handlers, with respect to the list of key handlers of the broadcaster, and a following handling by the broadcaster will not take place.

    **since**

        OOo 1.1.2

    See Also:
        `API XKeyHandler <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XKeyHandler.html>`_
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
            subscriber.addKeyHandler(self)

    # region XKeyListener
    def keyPressed(self, event: KeyEvent) -> None:
        """
        is invoked when a key has been pressed.
        """
        self._trigger_event("keyPressed", event)

    def keyReleased(self, event: KeyEvent) -> None:
        """
        is invoked when a key has been released.
        """
        self._trigger_event("keyReleased", event)

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

    # endregion XKeyListener
