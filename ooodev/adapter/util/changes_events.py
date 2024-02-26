from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.util import XChangesNotifier

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.loader import lo as mLo
from .changes_listener import ChangesListener

if TYPE_CHECKING:
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class ChangesEvents:
    """
    Class for managing Changes Events.

    This class is usually inherited by control classes that implement ``com.sun.star.util.XChangesListener``.

    .. versionadded:: 0.19.1
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: ChangesListener | None = None,
        subscriber: XChangesNotifier | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (ChangesListener | None, optional): Listener that is used to manage events.
            subscriber (XChangesNotifier, optional): An UNO object that implements the ``XChangesNotifier`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                mb = mLo.Lo.qi(XChangesNotifier, subscriber, True)
                mb.addChangesListener(self.__listener)
        else:
            self.__listener = ChangesListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_changes_occurred(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when something changes in the object.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.util.ChangesEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="changesOccurred")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("changesOccurred", cb)

    def add_event_changes_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_changes_occured(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="changesOccurred", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("changesOccurred", cb)

    def remove_event_changes_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_changes(self) -> ChangesListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
