from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.form import XApproveActionListener

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase
from ooodev.events.args.cancel_event_args import CancelEventArgs


if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.form import XApproveActionBroadcaster


class ApproveActionListener(AdapterBase, XApproveActionListener):
    """
    can be implemented to listen and probably veto actions to be performed on components.

    An example for an action would be the click of a CommandButton.

    See Also:
        `API XApproveActionListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1form_1_1XApproveActionListener.html>`_
    """

    def __init__(
        self, trigger_args: GenericArgs | None = None, subscriber: XApproveActionBroadcaster | None = None
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XApproveActionBroadcaster, optional): An UNO object that implements the ``XApproveActionBroadcaster`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)

        if subscriber:
            subscriber.addApproveActionListener(self)

    def approveAction(self, event: EventObject) -> bool:
        """
        Event is invoked when an action is preformed.
        If event is canceled then the action will be canceled.

        Args:
            event (EventObject): Event data for the event.

        Returns:
            bool: ``True`` if the action should be performed, ``False`` otherwise.

        Note:
            When ``approveAction`` event is invoked it will contain a :py:class:`~ooodev.events.args.cancel_event_args.CancelEventArgs`
            instance as the trigger event. When the event is triggered the ``CancelEventArgs.cancel`` can be set to ``True``
            to cancel the action. Also if canceled the ``CancelEventArgs.handled`` can be set to ``True`` to indicate that the action
            should be performed. The ``CancelEventArgs.event_data`` will contain the original ``com.sun.star.lang.EventObject``
            that triggered the update.
        """
        cancel_args = CancelEventArgs(self.__class__.__qualname__)
        cancel_args.event_data = event
        self._trigger_direct_event("approveAction", cancel_args)
        if cancel_args.cancel:
            if CancelEventArgs.handled:
                # if the cancel event was handled then we return True to indicate that the update should be performed
                return True
            return False
        return True

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
