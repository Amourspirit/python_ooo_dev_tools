from __future__ import annotations

from .focus_listener import FocusListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.utils.type_var import EventArgsCallbackT


class FocusEvents:
    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        self.__focus_listener = FocusListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_focus_gained(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window gains the keyboard focus.

        The callback ``EventArgs.event_data`` will contain a UNO ``FocusEvent`` struct.
        """
        self.__focus_listener.on("focusGained", cb)

    def add_event_focus_lost(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window loses the keyboard focus.

        The callback ``EventArgs.event_data`` will contain a UNO ``FocusEvent`` struct.
        """
        self.__focus_listener.on("focusLost", cb)

    def remove_event_focus_gained(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__focus_listener.off("focusGained", cb)

    def remove_event_focus_lost(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__focus_listener.off("focusLost", cb)

    # endregion Manage Events

    @property
    def events_listener_focus(self) -> FocusListener:
        """
        Returns listener
        """
        return self.__focus_listener
