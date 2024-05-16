from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.script.vba import XVBAScriptListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.script.vba import XVBACompatibility
    from com.sun.star.script.vba import VBAScriptEvent


class VBAScriptListener(AdapterBase, XVBAScriptListener):
    """
    VBA Script Listener.

    See Also:
        `API XVBAScriptListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1script_1_1vba_1_1XVBAScriptListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XVBACompatibility | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XVBACompatibility, optional): An UNO object that implements the ``XVBACompatibility`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)

        if subscriber:
            subscriber.addVBAScriptListener(self)

    def notifyVBAScriptEvent(self, event: VBAScriptEvent) -> None:
        """
        Event is invoked when the object is about to be reloaded.

        Components may use this to stop any other event processing related to the event source
        until they get the reloaded event.

        Args:
            event (VBAScriptEvent): Event data for the event.

        Returns:
            None:
        """
        self._trigger_event("notifyVBAScriptEvent", event)

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
