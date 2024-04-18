from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.frame import XTitleChangeListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.frame import TitleChangedEvent
    from com.sun.star.frame import XTitleChangeBroadcaster


class TitleChangeListener(AdapterBase, XTitleChangeListener):
    """
    Allows to receive notifications when the frame title changes.

    See Also:
        `API XTitleChangeListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XTitleChangeListener.html>`_
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        subscriber: XTitleChangeBroadcaster | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XTitleChangeBroadcaster, optional): An UNO object that implements the ``XTitleChangeBroadcaster`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber is not None:
            subscriber.addTitleChangeListener(self)

    def titleChanged(self, event: TitleChangedEvent) -> None:
        """
        Is invoked whenever the frame title has changed.
        """
        self._trigger_event("titleChanged", event)

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
