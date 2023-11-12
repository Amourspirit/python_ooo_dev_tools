from __future__ import annotations

from .adjustment_listener import AdjustmentListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class AdjustmentEvents:
    """
    Class for managing Adjustment Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XAdjustmentListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__adjustment_listener = AdjustmentListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_adjustment_value_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the adjustment has changed.

        The callback ``EventArgs.event_data`` will contain a UNO ``AdjustmentEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="adjustmentValueChanged")
            self.__callback(self, args)
        self.__adjustment_listener.on("adjustmentValueChanged", cb)

    def remove_event_adjustment_value_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="adjustmentValueChanged", is_add=False)
            self.__callback(self, args)
        self.__adjustment_listener.off("adjustmentValueChanged", cb)

    @property
    def events_listener_adjustment(self) -> AdjustmentListener:
        """
        Returns listener
        """
        return self.__adjustment_listener

    # endregion Manage Events
