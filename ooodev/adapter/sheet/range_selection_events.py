from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.sheet.range_selection_listener import RangeSelectionListener

if TYPE_CHECKING:
    from com.sun.star.sheet import XRangeSelection
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class RangeSelectionEvents:
    """
    Class for managing Activation Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: RangeSelectionListener | None = None,
        subscriber: XRangeSelection | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (RangeSelectionListener | None, optional): Listener that is used to manage events.
            subscriber (XRangeSelection, optional): An UNO object that implements the ``XRangeSelection`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addRangeSelectionListener(self.__listener)
        else:
            self.__listener = RangeSelectionListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_aborted(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when range selection is aborted.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.sheet.RangeSelectionEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="aborted")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("aborted", cb)

    def add_event_done(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when range selection is completed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.sheet.RangeSelectionEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="done")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("done", cb)

    def add_event_range_selection_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_aborted(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="aborted", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("aborted", cb)

    def remove_event_done(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="done", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("done", cb)

    def remove_event_range_selection_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_range_selection(self) -> RangeSelectionListener:
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
        source (Any): Expected to be an instance of RangeSelectionEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, RangeSelectionEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XRangeSelection", source.component)  # type: ignore
    comp.addRangeSelectionListener(source.events_listener_range_selection)
    event.remove_callback = True
