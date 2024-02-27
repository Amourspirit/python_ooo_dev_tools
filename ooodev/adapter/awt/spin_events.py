from __future__ import annotations

from typing import TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.awt.spin_listener import SpinListener

if TYPE_CHECKING:
    from com.sun.star.awt import XSpinField
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class SpinEvents:
    """
    Class for managing Spin Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XSpinListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: SpinListener | None = None,
        subscriber: XSpinField | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (SpinListener | None, optional): Listener that is used to manage events.
            subscriber (XSpinField, optional): An UNO object that implements the ``XSpinField`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addSpinListener(self.__listener)
        else:
            self.__listener = SpinListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_down(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the spin field is spun down.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.SpinEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="down")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("down", cb)

    def add_event_first(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the spin field is set to the lower value.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.SpinEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="first")
            self.__callback(self, args)
        self.__listener.on("first", cb)

    def add_event_last(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the spin field is set to the upper value.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.SpinEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="last")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("last", cb)

    def add_event_up(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the spin field is spun up.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.SpinEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="up")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("up", cb)

    def add_event_spin_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_down(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="down", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("down", cb)

    def remove_event_first(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="first", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("first", cb)

    def remove_event_up(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="up", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("up", cb)

    def remove_event_last(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="last", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("last", cb)

    def remove_event_spin_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_spin(self) -> SpinListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
