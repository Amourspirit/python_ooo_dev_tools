from __future__ import annotations

from typing import TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.frame.frame_action_listener import FrameActionListener

if TYPE_CHECKING:
    from com.sun.star.frame import XFrame
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class FrameActionEvents:
    """
    Class for managing FrameAction Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: FrameActionListener | None = None,
        subscriber: XFrame | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (FrameActionListener | None, optional): Listener that is used to manage events.
            subscriber (XFrame, optional): An UNO object that implements the ``XFrame`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener is not None:
            self.__listener = listener
            if subscriber:
                subscriber.addFrameActionListener(self.__listener)
        else:
            self.__listener = FrameActionListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_frame_action(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the master environment is finally terminated.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.frame.FrameActionEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="frameAction")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("frameAction", cb)

    def add_event_frame_action_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_frame_action(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="frameAction", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("frameAction", cb)

    def remove_event_frame_action_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_frame_action(self) -> FrameActionListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
