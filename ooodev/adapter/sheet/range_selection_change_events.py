from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.sheet.range_selection_change_listener import RangeSelectionChangeListener


if TYPE_CHECKING:
    from com.sun.star.sheet import XRangeSelection
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class RangeSelectionChangeEvents:
    """
    Class for managing Range Selection Change Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: RangeSelectionChangeListener | None = None,
        subscriber: XRangeSelection | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (RangeSelectionChangeListener | None, optional): Listener that is used to manage events.
            subscriber (XRangeSelection, optional): An UNO object that implements the ``XRangeSelection`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addRangeSelectionChangeListener(self.__listener)
        else:
            self.__listener = RangeSelectionChangeListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_descriptor_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the selected range is changed while range selection is active.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.sheet.RangeSelectionEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="descriptorChanged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("descriptorChanged", cb)

    def add_event_range_selection_change_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_descriptor_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="descriptorChanged", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("descriptorChanged", cb)

    def remove_event_range_selection_change_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_range_selection_change(self) -> RangeSelectionChangeListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
