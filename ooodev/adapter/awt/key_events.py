from __future__ import annotations

from .key_listener import KeyListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class KeyEvents:
    """
    Class for managing Key Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XKeyListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__key_listener = KeyListener(trigger_args=trigger_args)
        self.__name = "ooodev.adapter.awt.KeyEvents"

    # region Manage Events
    def add_event_key_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a key has been pressed.

        The callback ``EventArgs.event_data`` will contain a UNO ``KeyEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="keyPressed")
            self.__callback(self, args)
        self.__key_listener.on("keyPressed", cb)

    def add_event_key_released(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a key has been released.

        The callback ``EventArgs.event_data`` will contain a UNO ``KeyEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="keyReleased")
            self.__callback(self, args)
        self.__key_listener.on("keyReleased", cb)

    def remove_event_key_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="keyPressed", is_add=False)
            self.__callback(self, args)
        self.__key_listener.off("keyPressed", cb)

    def remove_event_key_released(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="keyReleased", is_add=False)
            self.__callback(self, args)
        self.__key_listener.off("keyReleased", cb)

    # endregion Manage Events

    @property
    def events_listener_key(self) -> KeyListener:
        """
        Returns listener
        """
        return self.__key_listener
