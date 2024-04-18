from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.awt.key_handler import KeyHandler

if TYPE_CHECKING:
    from com.sun.star.awt import XUserInputInterception
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class KeyHandlerEvents:
    """
    Class for managing Key Handler Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XKeyListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: KeyHandler | None = None,
        subscriber: XUserInputInterception | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (KeyHandler | None, optional): Listener that is used to manage events.
            subscriber (XUserInputInterception, optional): An UNO object that implements the ``XUserInputInterception`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addKeyHandler(self.__listener)
        else:
            self.__listener = KeyHandler(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_key_handler_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a key has been pressed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.KeyEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="keyPressed")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("keyPressed", cb)

    def add_event_key_handler_released(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a key has been released.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.KeyEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="keyReleased")
            self.__callback(self, args)
        self.__listener.on("keyReleased", cb)

    def add_event_key_handler_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_key_handler_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        # sourcery skip: class-extract-method
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="keyPressed", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("keyPressed", cb)

    def remove_event_key_handler_released(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="keyReleased", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("keyReleased", cb)

    def remove_event_key_handler_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_key_handler(self) -> KeyHandler:
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
        source (Any): Expected to be an instance of KeyHandlerEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, KeyHandlerEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XUserInputInterception", source.component)  # type: ignore
    comp.addKeyHandler(source.events_listener_key_handler)
    event.remove_callback = True
