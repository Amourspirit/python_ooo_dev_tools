from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.ucb import XCommandInfoChangeListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.ucb import XCommandInfoChangeNotifier
    from com.sun.star.ucb import CommandInfoChangeEvent
    from com.sun.star.lang import EventObject


class CommandInfoChangeListener(AdapterBase, XCommandInfoChangeListener):
    """
    A listener for events related to XContents.

    See Also:
        `API XCommandInfoChangeListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1ucb_1_1XContentEventListener.html>`_
    """

    def __init__(
        self, trigger_args: GenericArgs | None = None, subscriber: XCommandInfoChangeNotifier | None = None
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XCommandInfoChangeNotifier, optional): An UNO object that implements the ``com.sun.star.form.XContent`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)

        if subscriber:
            subscriber.addCommandInfoChangeListener(self)

    def commandInfoChange(self, event: CommandInfoChangeEvent) -> None:
        """
        Event is invoked when changes of a ``XCommandInfo`` shall be propagated.
        """
        self._trigger_event("commandInfoChange", event)

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
