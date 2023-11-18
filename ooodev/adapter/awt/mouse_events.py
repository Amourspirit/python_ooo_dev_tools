from __future__ import annotations

from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT
from .mouse_listener import MouseListener


class MouseEvents:
    """
    Class for managing Mouse Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XMouseListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__listener = MouseListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_mouse_entered(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the mouse enters a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseEntered")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("mouseEntered", cb)

    def add_event_mouse_exited(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the mouse exits a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseExited")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("mouseExited", cb)

    def add_event_mouse_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a mouse button has been pressed on a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mousePressed")
            self.__callback(self, args)
        self.__listener.on("mousePressed", cb)

    def add_event_mouse_released(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a mouse button has been released on a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseReleased")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("mouseReleased", cb)

    def remove_event_mouse_entered(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseEntered", is_add=False)
            self.__callback(self, args)
        self.__listener.off("mouseEntered", cb)

    def remove_event_mouse_exited(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseExited", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("mouseExited", cb)

    def remove_event_mouse_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.

        Event is invoked when a mouse button has been pressed on a window.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mousePressed", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("mousePressed", cb)

    def remove_event_mouse_released(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseReleased", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("mouseReleased", cb)

    # endregion Manage Events

    @property
    def events_listener_mouse(self) -> MouseListener:
        """
        Returns listener
        """
        return self.__listener
