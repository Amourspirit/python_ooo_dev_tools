from __future__ import annotations

from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT
from .mouse_motion_listener import MouseMotionListener


class MouseMotionEvents:
    """
    Class for managing Mouse Motion Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XMouseMotionListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__listener = MouseMotionListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_mouse_dragged(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a mouse button is pressed on a window and then dragged.

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseDragged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("mouseDragged", cb)

    def add_event_mouse_moved(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event Is invoked when the mouse pointer has been moved on a window (with no buttons down).

        The callback ``EventArgs.event_data`` will contain a UNO ``MouseEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseMoved")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("mouseMoved", cb)

    def remove_event_mouse_dragged(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseDragged", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("mouseDragged", cb)

    def remove_event_mouse_moved(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseMoved", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("mouseMoved", cb)

    # endregion Manage Events

    @property
    def events_listener_mouse_motion(self) -> MouseMotionListener:
        """
        Returns listener
        """
        return self.__listener
