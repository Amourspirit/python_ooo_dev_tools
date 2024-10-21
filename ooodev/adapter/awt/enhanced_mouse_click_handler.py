from __future__ import annotations
import contextlib
from typing import TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from com.sun.star.awt import XEnhancedMouseClickHandler
from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs


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
    @override
    def mousePressed(self, e: EnhancedMouseEvent) -> bool:
        """
        Event is invoked when a mouse button has been pressed on a window.

        Note:
            When ``mousePressed`` event is invoked it will contain a :py:class:`~ooodev.events.args.cancel_event_args.CancelEventArgs`
            instance as the trigger event. When the event is triggered the ``CancelEventArgs.cancel`` can be set to ``True``
            to cancel the action. Also if canceled the ``CancelEventArgs.handled`` can be set to ``True`` to indicate that the action
            should be performed. The ``CancelEventArgs.event_data`` will contain the original ``com.sun.star.awt.EnhancedMouseEvent``
            that triggered the update.
        """
        cancel_args = CancelEventArgs(self.__class__.__qualname__)
        cancel_args.event_data = e
        self._trigger_direct_event("mousePressed", cancel_args)
        if cancel_args.cancel:
            if CancelEventArgs.handled:
                # if the cancel event was handled then we return True to indicate that the update should be performed
                return True
            return False
        return True

    @override
    def mouseReleased(self, e: EnhancedMouseEvent) -> bool:
        """
        Event is invoked when a mouse button has been released on a window.

        Note:
            When ``mouseReleased`` event is invoked it will contain a :py:class:`~ooodev.events.args.cancel_event_args.CancelEventArgs`
            instance as the trigger event. When the event is triggered the ``CancelEventArgs.cancel`` can be set to ``True``
            to cancel the action. Also if canceled the ``CancelEventArgs.handled`` can be set to ``True`` to indicate that the action
            should be performed. The ``CancelEventArgs.event_data`` will contain the original ``com.sun.star.awt.EnhancedMouseEvent``
            that triggered the update.
        """
        cancel_args = CancelEventArgs(self.__class__.__qualname__)
        cancel_args.event_data = e
        self._trigger_direct_event("mouseReleased", cancel_args)
        if cancel_args.cancel:
            if CancelEventArgs.handled:
                # if the cancel event was handled then we return True to indicate that the update should be performed
                return True
            return False
        return True

    @override
    def disposing(self, Source: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", Source)

    # endregion XEnhancedMouseClickHandler
