from __future__ import annotations

from .window_listener import WindowListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class WindowEvents:
    """
    Class for managing Window Events.

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
        self.__name = "ooodev.adapter.awt.WindowEvents"
        self.__window_listener = WindowListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_window_hidden(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the window has been hidden.

        The callback ``EventArgs.event_data`` will contain a UNO ``WindowEvent`` struct.
        """
        self.__window_listener.on("windowHidden", cb)

    def add_event_window_moved(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the window has been moved.

        The callback ``EventArgs.event_data`` will contain a UNO ``WindowEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowMoved")
            self.__callback(self, args)
        self.__window_listener.on("windowMoved", cb)

    def add_event_window_resized(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the window has been resized.

        The callback ``EventArgs.event_data`` will contain a UNO ``WindowEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowResized")
            self.__callback(self, args)
        self.__window_listener.on("windowResized", cb)

    def add_event_window_shown(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the window has been shown.

        The callback ``EventArgs.event_data`` will contain a UNO ``WindowEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowShown")
            self.__callback(self, args)
        self.__window_listener.on("windowShown", cb)

    def remove_event_window_hidden(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowHidden")
            self.__callback(self, args)
        self.__window_listener.off("windowHidden", cb)

    def remove_event_window_moved(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowMoved", is_add=False)
            self.__callback(self, args)
        self.__window_listener.off("windowMoved", cb)

    def remove_event_window_resized(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowResized", is_add=False)
            self.__callback(self, args)
        self.__window_listener.off("windowResized", cb)

    def remove_event_window_shown(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowShown", is_add=False)
            self.__callback(self, args)
        self.__window_listener.off("windowShown", cb)

    # endregion Manage Events

    @property
    def events_listener_window(self) -> WindowListener:
        """
        Returns listener
        """
        return self.__window_listener
