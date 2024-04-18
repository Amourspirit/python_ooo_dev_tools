from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.frame import XFrameActionListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.frame import FrameActionEvent
    from com.sun.star.frame import XFrame


class FrameActionListener(AdapterBase, XFrameActionListener):
    """
    Has to be provided if an object wants to receive events when several things happen to components within frames of the desktop frame tree.

    E.g., you can receive events of instantiation/destruction and activation/deactivation of components.

    See Also:
        `API XFrameActionListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XFrameActionListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XFrame | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XFrame, optional): An UNO object that implements the ``XFrame`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber is not None:
            subscriber.addFrameActionListener(self)

    def frameAction(self, event: FrameActionEvent) -> None:
        """
        Is invoked whenever any action occurs to a component within a frame.
        """
        self._trigger_event("frameAction", event)

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
