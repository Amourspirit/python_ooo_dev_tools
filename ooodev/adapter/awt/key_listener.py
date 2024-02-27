from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.awt import XKeyListener
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import KeyEvent
    from com.sun.star.awt import XWindow


class KeyListener(AdapterBase, XKeyListener):
    """
    Makes it possible to receive keyboard events.

    See Also:
        `API XKeyListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XKeyListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XWindow | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XWindow, optional): An UNO object that implements the ``XWindow`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addKeyListener(self)

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
