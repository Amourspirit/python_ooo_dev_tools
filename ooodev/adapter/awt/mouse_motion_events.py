from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.awt.mouse_motion_listener import MouseMotionListener

if TYPE_CHECKING:
    from com.sun.star.presentation import XSlideShowView
    from com.sun.star.awt import XWindow
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class MouseMotionEvents:
    """
    Class for managing Mouse Motion Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XMouseMotionListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: MouseMotionListener | None = None,
        subscriber: XSlideShowView | XWindow | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (MouseMotionListener | None, optional): Listener that is used to manage events.
            subscriber (XSlideShowView, XWindow, optional): An UNO object that implements ``XSlideShowView`` or ``XWindow`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addMouseMotionListener(self.__listener)
        else:
            self.__listener = MouseMotionListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_mouse_dragged(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a mouse button is pressed on a window and then dragged.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.MouseEvent`` struct.
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

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.MouseEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseMoved")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("mouseMoved", cb)

    def add_event_mouse_motion_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_mouse_motion_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="disposing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("disposing", cb)

    # endregion Manage Events

    @property
    def events_listener_mouse_motion(self) -> MouseMotionListener:
        """
        Returns listener
        """
        return self.__listener


def on_lazy_cb(source: Any, event: ListenerEventArgs) -> None:
    """
    Callback that is invoked when an event is added or removed.

    This method is generally used to add the listener to the component in a lazy manner.
    This means this callback will only be called once in the lifetime of the component.

    Args:
        source (Any): Expected to be an instance of MouseMotionEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, MouseMotionEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XWindow", source.component)  # type: ignore
    comp.addMouseMotionListener(source.events_listener_mouse_motion)
    event.remove_callback = True
