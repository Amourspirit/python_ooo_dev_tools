from __future__ import annotations

from .mouse_listener import MouseListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.utils.type_var import EventArgsCallbackT


class MouseEvents:
    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        self.__mouse_listener = MouseListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_mouse_entered(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the mouse enters a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        self.__mouse_listener.on("mouseEntered", cb)

    def add_event_mouse_exited(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the mouse exits a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        self.__mouse_listener.on("mouseExited", cb)

    def add_event_mouse_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a mouse button has been pressed on a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        self.__mouse_listener.on("mousePressed", cb)

    def add_event_mouse_released(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a mouse button has been released on a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        self.__mouse_listener.on("mouseReleased", cb)

    def remove_event_mouse_entered(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__mouse_listener.off("mouseEntered", cb)

    def remove_event_mouse_exited(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.
        """
        self.__mouse_listener.off("mouseExited", cb)

    def remove_event_mouse_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.

        Event is invoked when a mouse button has been pressed on a window.
        """
        self.__mouse_listener.off("mousePressed", cb)

    def remove_event_mouse_released(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.
        """
        self.__mouse_listener.off("mouseReleased", cb)

    # endregion Manage Events

    @property
    def events_listener_mouse(self) -> MouseListener:
        """
        Returns listener
        """
        return self.__mouse_listener
