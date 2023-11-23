from __future__ import annotations
import contextlib
from typing import Any, TYPE_CHECKING

from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from .item_listener import ItemListener

if TYPE_CHECKING:
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class ItemEvents:
    """
    Class for managing Item Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XItemListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: ItemListener | None = None,
        subscriber: Any = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (ItemListener | None, optional): Listener that is used to manage events.
            subscriber (Any, optional): An UNO object that has a ``addItemListener()`` Method.
                If passed in then this instance listener is automatically added to it.
                Valid objects are: RadioButton, ComboBox, CheckBox,
                XItemEventBroadcaster or any other UNO object that has ``addItemListener()`` method.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                with contextlib.suppress(AttributeError):
                    subscriber.addItemListener(self.__listener)
        else:
            self.__listener = ItemListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_item_state_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when an item changes its state.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.ItemEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="itemStateChanged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("itemStateChanged", cb)

    def add_event_item_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_item_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="disposing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("disposing", cb)

    # endregion Manage Events

    @property
    def events_listener_item(self) -> ItemListener:
        """
        Returns listener
        """
        return self.__listener
