from __future__ import annotations

from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT
from .action_listener import ActionListener


class ActionEvents:
    """
    Class for managing Action Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XActionListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: ActionListener | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (ActionListener | None, optional): Listener that is used to manage events.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
        else:
            self.__listener = ActionListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

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
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("actionPerformed", cb)

    def remove_event_action_performed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="actionPerformed", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("actionPerformed", cb)

    @property
    def events_listener_action(self) -> ActionListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
