from __future__ import annotations

from .key_listener import KeyListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.utils.type_var import EventArgsCallbackT


class KeyEvents:
    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        self.__key_listener = KeyListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_key_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a key has been pressed.

        The callback ``EventArgs.event_data`` will contain a UNO ``KeyEvent`` struct.
        """
        self.__key_listener.on("keyPressed", cb)

    def add_event_key_released(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a key has been released.

        The callback ``EventArgs.event_data`` will contain a UNO ``KeyEvent`` struct.
        """
        self.__key_listener.on("keyReleased", cb)

    def remove_event_key_pressed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__key_listener.off("keyPressed", cb)

    def remove_event_key_released(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__key_listener.off("keyReleased", cb)

    # endregion Manage Events

    @property
    def events_listener_key(self) -> KeyListener:
        """
        Returns listener
        """
        return self.__key_listener
