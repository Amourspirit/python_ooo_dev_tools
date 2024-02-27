from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.awt.grid.grid_column_listener import GridColumnListener

if TYPE_CHECKING:
    from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class GridColumnEvents:
    """
    Class for managing Grid Column Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.grid.XGridColumnListener``.
    """

    def __init__(
        self,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: GridColumnListener | None = None,
    ) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
                This only applies if the listener is not passed.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
            listener (GridColumnListener | None, optional): Listener that is used to manage events.
        """
        self.__callback = cb
        if listener:
            self.__listener = listener
        else:
            self.__listener = GridColumnListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_column_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked after a column was modified.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.awt.grid.GridColumnEvent`` struct.
        """
        # sourcery skip: class-extract-method
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="columnChanged")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.on("columnChanged", cb)

    def add_event_grid_column_events_disposing(self, cb: EventArgsCallbackT) -> None:
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

    def remove_event_column_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="columnChanged", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        self.__listener.off("columnChanged", cb)

    def remove_event_grid_column_events_disposing(self, cb: EventArgsCallbackT) -> None:
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
    def events_listener_grid_column(self) -> GridColumnListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
