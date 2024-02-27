from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.chart.chart_data_change_event_listener import ChartDataChangeEventListener

if TYPE_CHECKING:
    from com.sun.star.chart import XChartData
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class ChartDataChangeEventEvents:
    """
    Class for managing Chart Data Change  Events.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: ChartDataChangeEventListener | None = None,
        subscriber: XChartData | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (ChartDataChangeEventListener | None, optional): Listener that is used to manage events.
            subscriber (XChartData, optional): An UNO object that implements the ``XChartData`` interface.
                If passed in then this instance listener is automatically added to it.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
            if subscriber:
                subscriber.addChartDataChangeEventListener(self.__listener)
        else:
            self.__listener = ChartDataChangeEventListener(trigger_args=trigger_args, subscriber=subscriber)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_chart_data_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when chart data changes in value or structure.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.chart.ChartDataChangeEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="chartDataChanged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("chartDataChanged", cb)

    def add_event_chart_data_change_event_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_chart_data_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="chartDataChanged", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("chartDataChanged", cb)

    def remove_event_chart_data_change_event_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_chart_data_change_event(self) -> ChartDataChangeEventListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
