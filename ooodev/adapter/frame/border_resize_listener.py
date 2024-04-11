from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.frame import XBorderResizeListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.event_args import EventArgs

if TYPE_CHECKING:
    from com.sun.star.frame import BorderWidths
    from com.sun.star.frame import XControllerBorder
    from com.sun.star.lang import EventObject
    from com.sun.star.uno import XInterface


class BorderResizeListener(AdapterBase, XBorderResizeListener):
    """
    Allows to listen to border resize events of a controller.

    See Also:
        `API XBorderResizeListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XBorderResizeListener.html>`_
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        subscriber: XControllerBorder | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            add_listener (bool, optional): If ``True`` listener is automatically added. Default ``True``.
            subscriber (XControllerBorder, optional): An UNO object that implements the ``XControllerBorder`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber is not None:
            subscriber.addBorderResizeListener(self)

    def borderWidthsChanged(self, obj: XInterface, new_size: BorderWidths) -> None:
        """
        notifies the listener that the controller's border widths have been changed.
        """
        event = EventArgs(self)
        event.event_data = {"obj": obj, "new_size": new_size}
        self._trigger_direct_event("borderWidthsChanged", event)

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
