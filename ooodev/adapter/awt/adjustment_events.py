from __future__ import annotations

from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT
from .adjustment_listener import AdjustmentListener


class AdjustmentEvents:
    """
    Class for managing Adjustment Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XAdjustmentListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: AdjustmentListener | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (AdjustmentListener | None, optional): Listener that is used to manage events.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
        else:
            self.__listener = AdjustmentListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_adjustment_value_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the adjustment has changed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.AdjustmentEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="adjustmentValueChanged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("adjustmentValueChanged", cb)

    def add_event_adjustment_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_adjustment_value_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="adjustmentValueChanged", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("adjustmentValueChanged", cb)

    def remove_event_adjustment_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_adjustment(self) -> AdjustmentListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
