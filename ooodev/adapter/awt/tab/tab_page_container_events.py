from __future__ import annotations

from .tab_page_container_listener import TabPageContainerListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class TabPageContainerEvents:
    """
    Class for managing Tab Container Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.tab.XTabPageContainerListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__listener = TabPageContainerListener(trigger_args=trigger_args)
        self.__name = "ooodev.adapter.awt.tab.TabPageContainerEvents"

    # region Manage Events
    def add_event_tab_page_activated(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked after a tab page was activated.

        The callback ``EventArgs.event_data`` will contain a UNO ``TabPageActivatedEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="tabPageActivated")
            self.__callback(self, args)
        self.__listener.on("tabPageActivated", cb)

    def remove_event_tab_page_activated(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="tabPageActivated", is_add=False)
            self.__callback(self, args)
        self.__listener.off("tabPageActivated", cb)

    @property
    def events_listener_tab_page_container(self) -> TabPageContainerListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
