from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.document.storage_change_listener import StorageChangeListener

if TYPE_CHECKING:
    from com.sun.star.document import XStorageBasedDocument
    from ooodev.utils.type_var import EventArgsCallbackT
    from ooodev.utils.type_var import ListenerEventCallbackT


class StorageChangeEventEvents:
    """
    Class for managing Storage Change Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: StorageChangeListener | None = None,
        subscriber: XStorageBasedDocument | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (StorageChangeListener | None, optional): Listener that is used to manage events.
            subscriber (XStorageBasedDocument, optional): An UNO object that implements the ``XStorageBasedDocument`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addStorageChangeListener(self.__listener)
        else:
            self.__listener = StorageChangeListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_notify_storage_change(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when range selection is completed.

        The callback ``EventArgs.event_data`` will contain a ``ooodev.events.args.event_args.EventArgs`` instance.

        The event, event data is a dictionary with the following keys:

        - ``document``: The document that is being switched to another storage.
        - ``storage``: The new storage that the document is being switched to.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="notifyStorageChange")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("notifyStorageChange", cb)

    def add_event_storage_change_event_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_notify_storage_change(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="notifyStorageChange", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("notifyStorageChange", cb)

    def remove_event_storage_change_event_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_storage_change_event(self) -> StorageChangeListener:
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
        source (Any): Expected to be an instance of StorageChangeEventEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, StorageChangeEventEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XStorageBasedDocument", source.component)  # type: ignore
    comp.addStorageChangeListener(source.events_listener_storage_change_event)
    event.remove_callback = True
