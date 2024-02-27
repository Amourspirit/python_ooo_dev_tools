from __future__ import annotations
import contextlib
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.awt import XEnhancedMouseClickHandler
from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.generic_args import GenericArgs


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import EnhancedMouseEvent
    from com.sun.star.sheet import XEnhancedMouseClickBroadcaster


# This class unlike most listeners does not end with the name Listeners


class EnhancedMouseClickHandler(AdapterBase, XEnhancedMouseClickHandler):
    """
    Makes it possible to receive enhanced events from the mouse.

    **since**

        OOo 2.0

    See Also:
        `API XEnhancedMouseClickHandler <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XEnhancedMouseClickHandler.html>`_
    """

    def __init__(
        self, trigger_args: GenericArgs | None = None, subscriber: XEnhancedMouseClickBroadcaster | None = None
    ) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XEnhancedMouseClickBroadcaster, optional): An UNO object that implements the ``XEnhancedMouseClickBroadcaster`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            with contextlib.suppress(AttributeError):
                subscriber.addEnhancedMouseClickHandler(self)

    # region XEnhancedMouseClickHandler
    def mousePressed(self, event: EnhancedMouseEvent) -> None:
        """
        Event is invoked when a mouse button has been pressed on a window.
        """
        self._trigger_event("mousePressed", event)

    def mouseReleased(self, event: EnhancedMouseEvent) -> None:
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

    # endregion XEnhancedMouseClickHandler
