from __future__ import annotations

from .property_change_listener import PropertyChangeListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.utils.type_var import EventArgsCallbackT


class PropertyChangeEvents:
    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        self.__property_change_listener = PropertyChangeListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_property_change(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event Is invoked when bound properties are changed.

        The callback ``EventArgs.event_data`` will contain a ``PropertyChangeEvent`` struct.
        """
        self.__property_change_listener.on("propertyChange", cb)

    def remove_event_property_change(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__property_change_listener.off("propertyChange", cb)

    # endregion Manage Events

    @property
    def events_listener_property_change(self) -> PropertyChangeListener:
        """
        Returns listener
        """
        return self.__property_change_listener
