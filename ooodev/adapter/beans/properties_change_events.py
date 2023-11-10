from __future__ import annotations

from .properties_change_listener import PropertiesChangeListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.utils.type_var import EventArgsCallbackT


class PropertiesChangeEvents:
    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        self.__properties_change_listener = PropertiesChangeListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_properties_change(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event Is invoked when bound properties are changed.

        The callback ``EventArgs.event_data`` will contain a tuple of ``PropertyChangeEvent`` objects.
        """
        self.__properties_change_listener.on("propertiesChange", cb)

    def remove_event_properties_change(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__properties_change_listener.off("propertiesChange", cb)

    # endregion Manage Events

    @property
    def events_listener_properties_change(self) -> PropertiesChangeEvents:
        """
        Returns listener
        """
        return self.__properties_change_listener
