from __future__ import annotations

from .grid_column_listener import GridColumnListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class GridColumnEvents:
    """
    Class for managing Grid Selection Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.grid.XGridColumnListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__listener = GridColumnListener(trigger_args=trigger_args)
        self.__name = "ooodev.adapter.awt.grid.GridColumnEvents"

    # region Manage Events
    def add_event_column_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked after a column was modified.

        The callback ``EventArgs.event_data`` will contain a UNO ``GridColumnEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="columnChanged")
            self.__callback(self, args)
        self.__listener.on("columnChanged", cb)

    def remove_event_column_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="columnChanged", is_add=False)
            self.__callback(self, args)
        self.__listener.off("columnChanged", cb)

    @property
    def events_listener_grid_column(self) -> GridColumnListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
