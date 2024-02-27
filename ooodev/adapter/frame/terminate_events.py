from __future__ import annotations

from typing import TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.loader import lo as mLo
from ooodev.adapter.frame.terminate_listener import TerminateListener

if TYPE_CHECKING:
    from com.sun.star.frame import XDesktop
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


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
        add_terminate_listener: bool = True,
        subscriber: XDesktop | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (TerminateListener | None, optional): Listener that is used to manage events.
            add_terminate_listener (bool, optional): If ``True`` listener is automatically added. Default ``True``.
            subscriber (XDesktop, optional): An UNO object that implements the ``XDesktop`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if add_terminate_listener:
                desktop = mLo.Lo.get_desktop()
                desktop.addTerminateListener(self.__listener)
            if subscriber:
                subscriber.addTerminateListener(self.__listener)
        else:
            self.__listener = TerminateListener(
                trigger_args=trigger_args, add_listener=add_terminate_listener, subscriber=subscriber
            )
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_notify_termination(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the master environment is finally terminated.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
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

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="queryTermination")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("queryTermination", cb)

    def add_event_terminate_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_terminate_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_terminate(self) -> TerminateListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
