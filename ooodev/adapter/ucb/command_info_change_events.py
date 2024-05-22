from __future__ import annotations

from typing import Any, cast, TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.ucb.command_info_change_listener import CommandInfoChangeListener

if TYPE_CHECKING:
    from com.sun.star.ucb import XCommandInfoChangeNotifier
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class CommandInfoChangeEvents:
    """
    Class for managing Command Info Change Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: CommandInfoChangeListener | None = None,
        subscriber: XCommandInfoChangeNotifier | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (CommandInfoChangeListener | None, optional): Listener that is used to manage events.
            subscriber (XCommandInfoChangeNotifier, optional): An UNO object that implements the ``com.sun.star.form.XCommandInfoChangeNotifier`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addCommandInfoChangeListener(self.__listener)
        else:
            self.__listener = CommandInfoChangeListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_command_info_change(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked the current column in a grid control changed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.ucb.CommandInfoChangeEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="commandInfoChange")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("commandInfoChange", cb)

    def add_event_command_info_change_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_command_info_change(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="commandInfoChange", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("commandInfoChange", cb)

    def remove_event_command_info_change_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_command_info_change(self) -> CommandInfoChangeListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events


def on_lazy_cb(source: Any, event: ListenerEventArgs) -> None:
    """
    Callback that is invoked when an event is added or removed.

    This method is generally used to add the listener to the component in a lazy manner.
    This means this callback will only be called once in the lifetime of the component.

    Args:
        source (Any): Expected to be an instance of CommandInfoChangeEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, CommandInfoChangeEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XCommandInfoChangeNotifier", source.component)  # type: ignore
    comp.addCommandInfoChangeListener(source.events_listener_command_info_change)
    event.remove_callback = True
