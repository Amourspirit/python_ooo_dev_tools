from __future__ import annotations

from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT
from .top_window_listener import TopWindowListener


class TopWindowEvents:
    """
    Class for managing Top Window Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XWindowListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__name = gUtil.Util.generate_random_string(10)
        self.__listener = TopWindowListener(trigger_args=trigger_args, add_listener=False)

    # region Manage Events
    def add_event_window_activated(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window is activated.

        The callback ``EventArgs.event_data`` will contain a UNO ``WindowEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowActivated")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("windowActivated", cb)

    def add_event_window_closed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window has been closed.

        The callback ``EventArgs.event_data`` will contain a UNO ``WindowEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowActivated")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("windowActivated", cb)

    def add_event_window_closing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window is in the process of being closed.

        The callback ``EventArgs.event_data`` will contain a UNO ``WindowEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowClosing")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("windowClosing", cb)

    def add_event_window_deactivated(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window is deactivated.

        The callback ``EventArgs.event_data`` will contain a UNO ``WindowEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowDeactivated")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("windowDeactivated", cb)

    def add_event_window_minimized(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window is iconified.

        The callback ``EventArgs.event_data`` will contain a UNO ``WindowEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowMinimized")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("windowMinimized", cb)

    def add_event_window_normalized(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window is deiconified.

        The callback ``EventArgs.event_data`` will contain a UNO ``WindowEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowNormalized")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("windowNormalized", cb)

    def add_event_window_opened(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window has been opened.

        The callback ``EventArgs.event_data`` will contain a UNO ``WindowEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowOpened")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("windowOpened", cb)

    def remove_event_window_activated(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowActivated", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("windowActivated", cb)

    def remove_event_window_closed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowClosed", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("windowClosed", cb)

    def remove_event_window_closing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowClosing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("windowClosing", cb)

    def remove_event_window_deactivated(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowDeactivated", is_add=False)
            self.__callback(self, args)
        self.__listener.off("windowDeactivated", cb)

    def remove_event_window_minimized(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowMinimized", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("windowMinimized", cb)

    def remove_event_window_normalized(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowNormalized", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("windowNormalized", cb)

    def remove_event_window_opened(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowOpened", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("windowOpened", cb)

    # endregion Manage Events

    @property
    def events_listener_top_window(self) -> TopWindowListener:
        """
        Returns listener
        """
        return self.__listener
