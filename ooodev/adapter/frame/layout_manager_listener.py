from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XLayoutManagerListener
from ooo.dyn.frame.layout_manager_events import LayoutManagerEventsEnum

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.event_args import EventArgs

if TYPE_CHECKING:
    from com.sun.star.frame import XLayoutManagerEventBroadcaster
    from com.sun.star.lang import EventObject


class LayoutManagerListener(AdapterBase, XLayoutManagerListener):
    """
    Makes it possible to receive events from a layout manager.

    Events are provided only for notification purposes only. All operations are handled internally by the layout manager component, so that GUI layout works properly regardless of whether a component registers such a listener or not.

    **since**

        OOo 2.0

    See Also:
        `API XLayoutManagerListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XLayoutManagerListener.html>`_
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        subscriber: XLayoutManagerEventBroadcaster | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            add_listener (bool, optional): If ``True`` listener is automatically added. Default ``True``.
            subscriber (XLayoutManagerEventBroadcaster, optional): An UNO object that implements the ``XLayoutManagerEventBroadcaster`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber is not None:
            subscriber.addLayoutManagerEventListener(self)

    def layoutEvent(self, source: EventObject, layout_event: int, info: Any) -> None:
        """
        Event is invoked when a layout manager has made a certain operation.
        """
        event = EventArgs(self)
        if layout_event is None:
            layout_enum = None
        else:
            layout_enum = LayoutManagerEventsEnum(layout_event)
        event.event_data = {"source": source, "layout_event": layout_enum, "info": info}
        self._trigger_direct_event("layoutEvent", event)

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
