from __future__ import annotations

from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT
from .terminate_listener import TerminateListener


class TerminateEvents:
    """
    Class for managing Terminate Events.

    This class is usually inherited by control classes that implement ``com.sun.star.frame.XTerminateListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: TerminateListener | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (TerminateListener | None, optional): Listener that is used to manage events.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
        else:
            self.__listener = TerminateListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_notify_termination(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the master environment is finally terminated.

        The callback ``EventArgs.event_data`` will contain a UNO ``EventObject`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="notifyTermination")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("notifyTermination", cb)

    def add_event_query_termination(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the master environment (e.g., desktop) is about to terminate.

        The callback ``EventArgs.event_data`` will contain a UNO ``EventObject`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="queryTermination")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("queryTermination", cb)

    def remove_event_notify_termination(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="notifyTermination", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("notifyTermination", cb)

    def remove_event_query_termination(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="queryTermination", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("queryTermination", cb)

    @property
    def events_listener_terminate(self) -> TerminateListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
