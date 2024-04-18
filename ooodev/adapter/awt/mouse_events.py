from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.awt.mouse_listener import MouseListener

if TYPE_CHECKING:
    from com.sun.star.presentation import XSlideShowView
    from com.sun.star.awt import XWindow
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class MouseEvents:
    """
    Class for managing Mouse Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XMouseListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: MouseListener | None = None,
        subscriber: XSlideShowView | XWindow | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (MouseListener | None, optional): Listener that is used to manage events.
            subscriber (XSlideShowView, XWindow, optional): An UNO object that implements ``XSlideShowView`` or ``XWindow`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addMouseListener(self.__listener)
        else:
            self.__listener = MouseListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_mouse_entered(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the mouse enters a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.MouseEvent`` struct.
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

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.MouseEvent`` struct.
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

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.MouseEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mousePressed")
            self.__callback(self, args)
        self.__listener.on("mousePressed", cb)

    def add_event_mouse_released(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a mouse button has been released on a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.MouseEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseReleased")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("mouseReleased", cb)

    def add_event_mouse_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_mouse_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_mouse(self) -> MouseListener:
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
        source (Any): Expected to be an instance of MouseEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, MouseEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XWindow", source.component)  # type: ignore
    comp.addMouseListener(source.events_listener_mouse)
    event.remove_callback = True
