from __future__ import annotations

from typing import TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.form.update_listener import UpdateListener

if TYPE_CHECKING:
    from ooodev.utils.type_var import EventArgsCallbackT, CancelEventArgsCallbackT, ListenerEventCallbackT
    from com.sun.star.form import XUpdateBroadcaster


class UpdateEvents:
    """
    Class for managing Update Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: UpdateListener | None = None,
        subscriber: XUpdateBroadcaster | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (UpdateListener | None, optional): Listener that is used to manage events.
            subscriber (XUpdateBroadcaster, optional): An UNO object that implements the ``XUpdateBroadcaster`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addUpdateListener(self.__listener)
        else:
            self.__listener = UpdateListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_approve_update(self, cb: CancelEventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked to check the current data. If event is canceled then the update will be canceled.

        The callback ``CancelEventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.

        Note:
            The callback event will be :py:class:`~ooodev.events.args.cancel_event_args.CancelEventArgs`.
            If the ``CancelEventArgs.cancel`` is set to ``True`` then the update will be canceled if the ``CancelEventArgs.handled``
            is set to ``True`` then the update will be performed.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="approveUpdate")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("approveUpdate", cb)

    def add_event_updated(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when an object has finished processing the updates and the data has been successfully written.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="updated")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("updated", cb)

    def add_event_update_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_updated(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="updated", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("updated", cb)

    def remove_event_approve_update(self, cb: CancelEventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="approveUpdate", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("approveUpdate", cb)

    def remove_event_update_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_update(self) -> UpdateListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
