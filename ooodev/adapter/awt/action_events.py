from __future__ import annotations

from .action_listener import ActionListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.utils.type_var import EventArgsCallbackT


class ActionEvents:
    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        self.__action_listener = ActionListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_action_performed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when an action is performed.

        The callback ``EventArgs.event_data`` will contain a UNO ``ActionEvent`` struct.
        """
        self.__action_listener.on("actionPerformed", cb)

    def remove_event_action_performed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__action_listener.off("actionPerformed", cb)

    @property
    def events_listener_action(self) -> ActionListener:
        """
        Returns listener
        """
        return self.__action_listener

    # endregion Manage Events
