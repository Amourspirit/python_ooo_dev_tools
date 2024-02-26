from __future__ import annotations

from typing import TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from .reset_listener import ResetListener

if TYPE_CHECKING:
    from com.sun.star.form import XReset
    from ooodev.utils.type_var import EventArgsCallbackT, CancelEventArgsCallbackT, ListenerEventCallbackT


class ResetEvents:
    """
    Class for managing Reset Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: ResetListener | None = None,
        subscriber: XReset | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (ResetListener | None, optional): Listener that is used to manage events.
            subscriber (XReset, optional): An UNO object that implements the ``XReset`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addResetListener(self.__listener)
        else:
            self.__listener = ResetListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_approve_reset(self, cb: CancelEventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked is invoked before a component is reset. If event is canceled then the reset will be canceled.

        The callback ``CancelEventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.

        Note:
            The callback event will be :py:class:`~ooodev.events.args.cancel_event_args.CancelEventArgs`.
            If the ``CancelEventArgs.cancel`` is set to ``True`` then the reset will be canceled if the ``CancelEventArgs.handled``
            is set to ``True`` then the reset will be performed.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="approveReset")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("approveReset", cb)

    def add_event_resetted(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a component has been reset.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="resetted")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("resetted", cb)

    def add_event_reset_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_approve_reset(self, cb: CancelEventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="approveReset", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("approveReset", cb)

    def remove_event_resetted(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="resetted", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("resetted", cb)

    def remove_event_reset_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_reset(self) -> ResetListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
