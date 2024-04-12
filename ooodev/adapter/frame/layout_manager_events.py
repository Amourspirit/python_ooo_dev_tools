from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.frame.layout_manager_listener import LayoutManagerListener

if TYPE_CHECKING:
    from com.sun.star.frame import XLayoutManagerEventBroadcaster
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class LayoutManagerEvents:
    """
    Class for managing Border resize Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: LayoutManagerListener | None = None,
        subscriber: XLayoutManagerEventBroadcaster | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (LayoutManagerListener | None, optional): Listener that is used to manage events.
            subscriber (XLayoutManagerEventBroadcaster, optional): An UNO object that implements the ``XBorderResizeListener`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener is not None:
            self.__listener = listener
            if subscriber:
                subscriber.addLayoutManagerEventListener(self.__listener)
        else:
            self.__listener = LayoutManagerListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_layout_event(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a layout manager has made a certain operation.

        The callback ``EventArgs.event_data`` will contain a dictionary with keys ``source``, ``layout_event`` and ``info``.
        The ``source`` will be an ``com.sun.star.lang.EventObject`` and ``info`` is ``Any``.
        ``layout_event`` will be an enum value of ``LayoutManagerEventsEnum``.

        Hint:
            - ``LayoutManagerEventsEnum`` is an enum that can be imported from ``ooo.dyn.frame.layout_manager_events``.

        Note:
            ``LayoutManagerEventsEnum`` is an enum that is dynamically generated from ``LayoutManagerEvents`` constants.

        See Also:
            - `API LayoutManagerEvents <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1frame_1_1LayoutManagerEvents.html>`__
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="layoutEvent")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("layoutEvent", cb)

    def add_event_layout_manager_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_layout_event(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="layoutEvent", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("layoutEvent", cb)

    def remove_event_layout_manager_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_layout_manager(self) -> LayoutManagerListener:
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
        source (Any): Expected to be an instance of LayoutManagerEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # will only ever fire once
    if not isinstance(source, LayoutManagerEvents):
        return
    if not hasattr(source, "component"):
        return

    comp = cast("XLayoutManagerEventBroadcaster", source.component)  # type: ignore
    comp.addLayoutManagerEventListener(source.events_listener_layout_manager)
    event.remove_callback = True
