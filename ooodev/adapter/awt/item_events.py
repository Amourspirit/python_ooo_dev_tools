from __future__ import annotations
from typing import cast

from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT
from .item_listener import ItemListener


class ItemEvents:
    """
    Class for managing Item Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XItemListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__listener = ItemListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_item_state_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when an item changes its state.

        The callback ``EventArgs.event_data`` will contain a UNO ``ItemEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="itemStateChanged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("itemStateChanged", cb)

    def remove_event_item_state_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="itemStateChanged", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("itemStateChanged", cb)

    # endregion Manage Events

    @property
    def events_listener_item(self) -> ItemListener:
        """
        Returns listener
        """
        return self.__listener
