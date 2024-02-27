from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.document.document_event_listener import DocumentEventListener

if TYPE_CHECKING:
    from com.sun.star.document import XDocumentEventBroadcaster
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class DocumentEventEvents:
    """
    Class for managing Activation Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: DocumentEventListener | None = None,
        subscriber: XDocumentEventBroadcaster | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (DocumentEventListener | None, optional): Listener that is used to manage events.
            subscriber (XDocumentEventBroadcaster, optional): An UNO object that implements the ``XDocumentEventBroadcaster`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addDocumentEventListener(self.__listener)
        else:
            self.__listener = DocumentEventListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_document_event_occurred(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when range selection is completed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.document.DocumentEvent`` struct.
        """
        # renamed in version 0.19.1, was named add_event_document_event_occured, alias added below
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="documentEventOccured")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("documentEventOccured", cb)

    # alias for add_event_document_event_occurred, added in version 0.19.1
    add_event_document_event_occured = add_event_document_event_occurred

    def add_event_document_event_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_document_event_occurred(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        # renamed in version 0.19.1, was named remove_event_document_event_occured, alias added below
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="documentEventOccured", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("documentEventOccured", cb)

    # alias for remove_event_document_event_occurred, added in version 0.19.1
    remove_event_document_event_occured = remove_event_document_event_occurred

    def remove_event_document_event_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_document_event(self) -> DocumentEventListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
