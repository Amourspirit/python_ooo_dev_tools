from __future__ import annotations
import contextlib
from typing import TYPE_CHECKING
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from .enhanced_mouse_click_handler import EnhancedMouseClickHandler

if TYPE_CHECKING:
    from com.sun.star.sheet import XEnhancedMouseClickBroadcaster
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class EnhancedMouseClickEvents:
    """
    Class for managing Enhanced Mouse Click Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: EnhancedMouseClickHandler | None = None,
        subscriber: XEnhancedMouseClickBroadcaster | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (EnhancedMouseClickHandler | None, optional): Listener that is used to manage events.
            subscriber (XEnhancedMouseClickBroadcaster, optional): An UNO object that implements the ``XEnhancedMouseClickBroadcaster`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                # several object such as Scrollbar and SpinValue
                # There is no common interface for this, so we have to try them all.
                with contextlib.suppress(AttributeError):
                    subscriber.addEnhancedMouseClickHandler(self.__listener)
        else:
            self.__listener = EnhancedMouseClickHandler(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_mouse_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a mouse button has been pressed on a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.EnhancedMouseEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mousePressed")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("mousePressed", cb)

    def add_event_mouse_released(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a mouse button has been released on a window.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.EnhancedMouseEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseReleased")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("mouseReleased", cb)

    def add_event_enhanced_mouse_click_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_mouse_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mousePressed", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("mousePressed", cb)

    def remove_event_mouse_released(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="mouseReleased", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("mouseReleased", cb)

    def remove_event_enhanced_mouse_click_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_enhanced_mouse_click(self) -> EnhancedMouseClickHandler:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
