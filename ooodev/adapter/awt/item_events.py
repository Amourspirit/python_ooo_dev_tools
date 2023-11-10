from __future__ import annotations

from .item_listener import ItemListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.utils.type_var import EventArgsCallbackT


class ItemEvents:
    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        self.__item_listener = ItemListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_item_state_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when an item changes its state.

        The callback ``EventArgs.event_data`` will contain a UNO ``ItemEvent`` struct.
        """
        self.__item_listener.on("itemStateChanged", cb)

    def remove_event_item_state_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__item_listener.off("itemStateChanged", cb)

    # endregion Manage Events

    @property
    def events_listener_item(self) -> ItemListener:
        """
        Returns listener
        """
        return self.__item_listener
