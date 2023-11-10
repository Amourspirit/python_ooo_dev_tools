from __future__ import annotations

from .mouse_motion_listener import MouseMotionListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.utils.type_var import EventArgsCallbackT


class MouseMotionEvents:
    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        self.__mouse_motion_listener = MouseMotionListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_mouse_dragged(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a mouse button is pressed on a window and then dragged.

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        self.__mouse_motion_listener.on("mouseDragged", cb)

    def add_event_mouse_moved(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event Is invoked when the mouse pointer has been moved on a window (with no buttons down).

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        self.__mouse_motion_listener.on("mouseMoved", cb)

    def remove_event_mouse_dragged(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__mouse_motion_listener.off("mouseDragged", cb)

    def remove_event_mouse_moved(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.
        """
        self.__mouse_motion_listener.off("mouseMoved", cb)

    # endregion Manage Events

    @property
    def events_listener_mouse_motion(self) -> MouseMotionListener:
        """
        Returns listener
        """
        return self.__mouse_motion_listener
