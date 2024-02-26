from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.form import XChangeListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.form import XChangeBroadcaster


class ChangeListener(AdapterBase, XChangeListener):
    """
    Is the listener interface for receiving notifications about data changes.

    The concrete semantics of a change (i.e. the conditions for when a change event is fired) must be specified in the description
    of the service broadcasting the change.

    See Also:
        `API XChangeListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1form_1_1XChangeListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XChangeBroadcaster | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XChangeBroadcaster, optional): An UNO object that implements the ``com.sun.star.form.XChangeBroadcaster`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)

        if subscriber:
            subscriber.addChangeListener(self)

    def changed(self, event: EventObject) -> None:
        """
        Event is invoked when the data of a component has been changed.

        Args:
            event (EventObject): Event data for the event.

        Returns:
            None:
        """
        self._trigger_event("changed", event)

    def disposing(self, event: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.

        Args:
            event (EventObject): Event data for the event.

        Returns:
            None:
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", event)
