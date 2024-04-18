from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.awt.focus_listener import FocusListener

if TYPE_CHECKING:
    from com.sun.star.awt import XExtendedToolkit
    from com.sun.star.awt import XWindow
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class FocusEvents:
    """
    Class for managing Focus Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XFocusListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: FocusListener | None = None,
        subscriber: XExtendedToolkit | XWindow | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (FocusListener | None, optional): Listener that is used to manage events.
            subscriber (XExtendedToolkit, XWindow, optional): An UNO object that implements the ``XExtendedToolkit`` or ``XWindow`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addFocusListener(self.__listener)
        else:
            self.__listener = FocusListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_focus_gained(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window gains the keyboard focus.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.FocusEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="focusGained")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("focusGained", cb)

    def add_event_focus_lost(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window loses the keyboard focus.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.FocusEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__class__.__qualname__, trigger_name="focusLost")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("focusLost", cb)

    def add_event_focus_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_focus_gained(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="focusGained", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("focusGained", cb)

    def remove_event_focus_lost(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="focusLost", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("focusLost", cb)

    def remove_event_focus_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_focus(self) -> FocusListener:
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
        source (Any): Expected to be an instance of FocusEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, FocusEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XExtendedToolkit", source.component)  # type: ignore
    comp.addFocusListener(source.events_listener_focus)
    event.remove_callback = True
