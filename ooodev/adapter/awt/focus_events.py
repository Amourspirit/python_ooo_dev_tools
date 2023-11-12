from __future__ import annotations
from .focus_listener import FocusListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT

# from ooodev.events.args.event_args


class FocusEvents:
    """
    Class for managing Focus Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XFocusListener``.
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
        self.__focus_listener = FocusListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_focus_gained(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window gains the keyboard focus.

        The callback ``EventArgs.event_data`` will contain a UNO ``FocusEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="focusGained")
            self.__callback(self, args)
        self.__focus_listener.on("focusGained", cb)

    def add_event_focus_lost(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a window loses the keyboard focus.

        The callback ``EventArgs.event_data`` will contain a UNO ``FocusEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__class__.__qualname__, trigger_name="focusLost")
            self.__callback(self, args)
        self.__focus_listener.on("focusLost", cb)

    def remove_event_focus_gained(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="focusGained", is_add=False)
            self.__callback(self, args)
        self.__focus_listener.off("focusGained", cb)

    def remove_event_focus_lost(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="focusLost", is_add=False)
            self.__callback(self, args)
        self.__focus_listener.off("focusLost", cb)

    # endregion Manage Events

    @property
    def events_listener_focus(self) -> FocusListener:
        """
        Returns listener
        """
        return self.__focus_listener
