from __future__ import annotations

from typing import TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.form.approve_action_listener import ApproveActionListener

if TYPE_CHECKING:
    from com.sun.star.form import XApproveActionBroadcaster
    from ooodev.utils.type_var import EventArgsCallbackT, CancelEventArgsCallbackT, ListenerEventCallbackT


class ApproveActionEvents:
    """
    Class for managing Approve Action Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: ApproveActionListener | None = None,
        subscriber: XApproveActionBroadcaster | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (ApproveActionListener | None, optional): Listener that is used to manage events.
            subscriber (XApproveActionBroadcaster, optional): An UNO object that implements the ``XApproveActionBroadcaster`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addApproveActionListener(self.__listener)
        else:
            self.__listener = ApproveActionListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_approve_action(self, cb: CancelEventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked is invoked when an action is performed.
        If event is canceled then the action will be canceled.

        The callback ``CancelEventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.

        Note:
            The callback event will be :py:class:`~ooodev.events.args.cancel_event_args.CancelEventArgs`.
            If the ``CancelEventArgs.cancel`` is set to ``True`` then the action will be canceled if the ``CancelEventArgs.handled``
            is set to ``True`` then the action will be performed.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="approveAction")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("approveAction", cb)

    def add_event_approve_action_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the broadcaster is about to be disposed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """

        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="disposing")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("disposing", cb)

    def remove_event_approve_action(self, cb: CancelEventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="approveAction", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("approveAction", cb)

    def remove_event_approve_action_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="disposing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("disposing", cb)

    @property
    def events_listener_approve_action(self) -> ApproveActionListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
