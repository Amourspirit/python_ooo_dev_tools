from __future__ import annotations

from .text_listener import TextListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class TextEvents:
    """
    Class for managing Text Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XTextListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__name = gUtil.Util.generate_random_string(10)
        self.__text_listener = TextListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_text_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the text has changed.

        The callback ``EventArgs.event_data`` will contain a UNO ``TextEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="textChanged")
            self.__callback(self, args)
        self.__text_listener.on("textChanged", cb)

    def remove_event_text_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="textChanged", is_add=False)
            self.__callback(self, args)
        self.__text_listener.off("textChanged", cb)

    @property
    def events_listener_text(self) -> TextListener:
        """
        Returns listener
        """
        return self.__text_listener

    # endregion Manage Events
