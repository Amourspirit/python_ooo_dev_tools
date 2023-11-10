from __future__ import annotations

from .text_listener import TextListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.utils.type_var import EventArgsCallbackT


class TextEvents:
    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        self.__text_listener = TextListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_text_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the text has changed.

        The callback ``EventArgs.event_data`` will contain a UNO ``TextEvent`` struct.
        """
        self.__text_listener.on("textChanged", cb)

    def remove_event_text_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__text_listener.off("textChanged", cb)

    @property
    def events_listener_text(self) -> TextListener:
        """
        Returns listener
        """
        return self.__text_listener

    # endregion Manage Events
