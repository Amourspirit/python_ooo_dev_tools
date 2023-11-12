from __future__ import annotations

from .grid_data_listener import GridDataListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class GridDataEvents:
    """
    Class for managing Grid Data Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.grid.XGridDataListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__listener = GridDataListener(trigger_args=trigger_args)
        self.__name = "ooodev.adapter.awt.grid.GridDataEvents"

    # region Manage Events
    def add_event_data_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when existing data in a grid control's data model has been modified.

        The callback ``EventArgs.event_data`` will contain a UNO ``GridDataEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="dataChanged")
            self.__callback(self, args)
        self.__listener.on("dataChanged", cb)

    def add_event_row_heading_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        EEvent is invoked when the title of one or more rows changed.

        The callback ``EventArgs.event_data`` will contain a UNO ``GridDataEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="rowHeadingChanged")
            self.__callback(self, args)
        self.__listener.on("rowHeadingChanged", cb)

    def add_event_rows_inserted(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when one or more rows of data have been inserted into a grid control's data model.is invoked when the title of one or more rows changed.

        The callback ``EventArgs.event_data`` will contain a UNO ``GridDataEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="rowsInserted")
            self.__callback(self, args)
        self.__listener.on("rowsInserted", cb)

    def add_event_rows_removed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is is invoked when one or more rows of data have been removed from a grid control's data model.

        The callback ``EventArgs.event_data`` will contain a UNO ``GridDataEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="rowsRemoved")
            self.__callback(self, args)
        self.__listener.on("rowsRemoved", cb)

    def remove_event_data_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="dataChanged", is_add=False)
            self.__callback(self, args)
        self.__listener.off("dataChanged", cb)

    def remove_event_row_heading_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="rowHeadingChanged", is_add=False)
            self.__callback(self, args)
        self.__listener.off("rowHeadingChanged", cb)

    def remove_event_rows_inserted(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="rowsInserted", is_add=False)
            self.__callback(self, args)
        self.__listener.off("rowsInserted", cb)

    def remove_event_rows_removed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="rowsRemoved", is_add=False)
            self.__callback(self, args)
        self.__listener.off("rowsRemoved", cb)

    @property
    def events_listener_grid_data(self) -> GridDataListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
