from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.awt import XTextListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.awt import TextEvent
    from com.sun.star.awt import XTextComponent


class TextListener(AdapterBase, XTextListener):
    """
    Makes it possible to receive text change events.

    See Also:
        `API XTextListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XTextListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XTextComponent | None = None) -> None:
        """
        Constructor:

        Arguments:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            subscriber (XTextComponent, optional): An UNO object that implements the ``XTextComponent`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        if subscriber:
            subscriber.addTextListener(self)

    # region XTextListener

    def textChanged(self, event: TextEvent) -> None:
        """
        Event is invoked when the text has changed.
        """
        self._trigger_event("textChanged", event)

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

    # endregion XTextListener
