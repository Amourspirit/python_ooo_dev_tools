from __future__ import annotations

from .spin_listener import SpinListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class SpinEvents:
    """
    Class for managing Spin Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XSpinListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__name = gUtil.Util.generate_random_string(10)
        self.__spin_listener = SpinListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_down(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the spin field is spun down.

        The callback ``EventArgs.event_data`` will contain a UNO ``SpinEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="down")
            self.__callback(self, args)
        self.__spin_listener.on("down", cb)

    def add_event_first(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the spin field is set to the lower value.

        The callback ``EventArgs.event_data`` will contain a UNO ``SpinEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="first")
            self.__callback(self, args)
        self.__spin_listener.on("first", cb)

    def add_event_last(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the spin field is set to the upper value.

        The callback ``EventArgs.event_data`` will contain a UNO ``SpinEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="last")
            self.__callback(self, args)
        self.__spin_listener.on("last", cb)

    def add_event_up(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the spin field is spun up.

        The callback ``EventArgs.event_data`` will contain a UNO ``SpinEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="up")
            self.__callback(self, args)
        self.__spin_listener.on("up", cb)

    def remove_event_down(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="down", is_add=False)
            self.__callback(self, args)
        self.__spin_listener.off("down", cb)

    def remove_event_first(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="first", is_add=False)
            self.__callback(self, args)
        self.__spin_listener.off("first", cb)

    def remove_event_up(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="up", is_add=False)
            self.__callback(self, args)
        self.__spin_listener.off("up", cb)

    def remove_event_last(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="last", is_add=False)
            self.__callback(self, args)
        self.__spin_listener.off("last", cb)

    @property
    def events_listener_spin(self) -> SpinListener:
        """
        Returns listener
        """
        return self.__spin_listener

    # endregion Manage Events
