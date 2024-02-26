from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.awt.tab.tab_page_container_listener import TabPageContainerListener

if TYPE_CHECKING:
    from com.sun.star.awt.tab import XTabPageContainer
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class TabPageContainerEvents:
    """
    Class for managing Tab Container Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.tab.XTabPageContainerListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: TabPageContainerListener | None = None,
        subscriber: XTabPageContainer | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (GridColumnListener | None, optional): Listener that is used to manage events.
            subscriber (XTabPageContainer, optional): An UNO object that implements the ``XTabPageContainer`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addTabPageContainerListener(self.__listener)
        else:
            self.__listener = TabPageContainerListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_tab_page_activated(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked after a tab page was activated.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.tab.TabPageActivatedEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="tabPageActivated")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("tabPageActivated", cb)

    def add_event_tab_page_container_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_tab_page_activated(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="tabPageActivated", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("tabPageActivated", cb)

    def remove_event_tab_page_container_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_tab_page_container(self) -> TabPageContainerListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
