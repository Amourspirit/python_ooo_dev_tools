from __future__ import annotations

from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT
from .container_listener import ContainerListener


class ContainerEvents:
    """
    Class for managing Container Events.

    This class is usually inherited by control classes that implement ``com.sun.star.container.XContainerListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: ContainerListener | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (ContainerListener | None, optional): Listener that is used to manage events.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
        else:
            self.__listener = ContainerListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_element_inserted(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a container has inserted an element.

        The callback ``EventArgs.event_data`` will contain a UNO ``ContainerEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="elementInserted")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("elementInserted", cb)

    def add_event_element_removed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a container has removed an element.

        The callback ``EventArgs.event_data`` will contain a UNO ``ContainerEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="elementRemoved")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("elementRemoved", cb)

    def add_event_element_replaced(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when a container has replaced an element.

        The callback ``EventArgs.event_data`` will contain a UNO ``ContainerEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="elementReplaced")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("elementReplaced", cb)

    def remove_event_element_inserted(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="elementInserted", is_add=False)
            self.__callback(self, args)
        self.__listener.off("elementInserted", cb)

    def remove_event_element_removed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="elementRemoved", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("elementRemoved", cb)

    def remove_event_element_replaced(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="elementReplaced", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("elementReplaced", cb)

    @property
    def events_listener_container(self) -> ContainerListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
