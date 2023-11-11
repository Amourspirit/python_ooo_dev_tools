from __future__ import annotations

from .action_listener import ActionListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class ActionEvents:
    """
    Class for managing Action Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XActionListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__action_listener = ActionListener(trigger_args=trigger_args)
        self.__name = "ooodev.adapter.awt.ActionEvents"

    # region Manage Events
    def add_event_action_performed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when an action is performed.

        The callback ``EventArgs.event_data`` will contain a UNO ``ActionEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="actionPerformed")
            self.__callback(self, args)
        self.__action_listener.on("actionPerformed", cb)

    def remove_event_action_performed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="actionPerformed", is_add=False)
            self.__callback(self, args)
        self.__action_listener.off("actionPerformed", cb)

    @property
    def events_listener_action(self) -> ActionListener:
        """
        Returns listener
        """
        return self.__action_listener

    # endregion Manage Events
