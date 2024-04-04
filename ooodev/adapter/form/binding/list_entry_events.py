from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.form.binding.list_entry_listener import ListEntryListener

if TYPE_CHECKING:
    from com.sun.star.form.binding import XListEntrySource
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class ListEntryEvents:
    """
    Class for managing List Entry Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: ListEntryListener | None = None,
        subscriber: XListEntrySource | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (ListEntryListener | None, optional): Listener that is used to manage events.
            subscriber (XListEntrySource, XWindow, optional): An UNO object that implements the ``XExtendedToolkit`` or ``XWindow`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addListEntryListener(self.__listener)
        else:
            self.__listener = ListEntryListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    def add_event_all_entries_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Notifies the listener that all entries of the list have changed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.

        The listener should retrieve the complete new list by calling the
        ``XListEntrySource.getAllListEntries()`` method of the event source
        (which is denoted by com.sun.star.lang.EventObject.Source).
        """

        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="allEntriesChanged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("allEntriesChanged", cb)

    # region Manage Events
    def add_event_entry_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Notifies the listener that a single entry in the list has change.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.form.binding.ListEntryEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="entryChanged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("entryChanged", cb)

    def add_event_entry_range_inserted(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Notifies the listener that a range of entries has been inserted into the list.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.form.binding.ListEntryEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="entryRangeInserted")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("entryRangeInserted", cb)

    def add_event_entry_range_removed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Notifies the listener that a range of entries has been removed from the list.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.form.binding.ListEntryEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="entryRangeRemoved")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("entryRangeRemoved", cb)

    def add_event_list_entry_listener_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_all_entries_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="allEntriesChanged", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("allEntriesChanged", cb)

    def remove_event_entry_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="entryChanged", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("entryChanged", cb)

    def remove_event_entry_range_inserted(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="entryRangeInserted", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("entryRangeInserted", cb)

    def remove_event_entry_range_removed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="entryRangeRemoved", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("entryRangeRemoved", cb)

    def remove_event_list_entry_listener_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_list_entry(self) -> ListEntryListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events


def on_lazy_cb(source: Any, event: ListenerEventArgs) -> None:
    """
    Callback that is invoked when an event is added or removed.

    This method is generally used to add the listener to the component in a lazy manner.
    This means this callback will only be called once in the lifetime of the component.

    Args:
        source (Any): Expected to be an instance of ListEntryEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, ListEntryEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XListEntrySource", source.component)  # type: ignore
    comp.addListEntryListener(source.events_listener_list_entry)
    event.remove_callback = True
