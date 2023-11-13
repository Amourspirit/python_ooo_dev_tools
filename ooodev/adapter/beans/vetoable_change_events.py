from __future__ import annotations

from .vetoable_change_listener import VetoableChangeListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class VetoableChangeEvents:
    """Class for managing Vetoable Change Events."""

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__name = "ooodev.adapter.beans.VetoableChangeEvents"
        self.__vetoable_change_listener = VetoableChangeListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_vetoable_change(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event Is invoked when bound properties are changed.

        The callback ``EventArgs.event_data`` will contain a ``PropertyChangeEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="vetoableChange")
            self.__callback(self, args)
        self.__vetoable_change_listener.on("vetoableChange", cb)

    def remove_event_vetoable_change(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="vetoableChange", is_add=False)
            self.__callback(self, args)
        self.__vetoable_change_listener.off("vetoableChange", cb)

    # endregion Manage Events

    @property
    def events_listener_vetoable_change(self) -> VetoableChangeListener:
        """
        Returns listener
        """
        return self.__vetoable_change_listener
