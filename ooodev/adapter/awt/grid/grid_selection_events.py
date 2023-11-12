from __future__ import annotations

from .grid_selection_listener import GridSelectionListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class GridSelectionEvents:
    """
    Class for managing Grid Selection Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.grid.XGridSelectionListener``.
    """

    def __init__(self, trigger_args: GenericArgs | None = None, cb: ListenerEventCallbackT | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__listener = GridSelectionListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_selection_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked after a selection was changed.

        The callback ``EventArgs.event_data`` will contain a UNO ``GridSelectionEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="selectionChanged")
            self.__callback(self, args)
        self.__listener.on("selectionChanged", cb)

    def remove_event_selection_changed(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="selectionChanged", is_add=False)
            self.__callback(self, args)
        self.__listener.off("selectionChanged", cb)

    @property
    def events_listener_grid_selection(self) -> GridSelectionListener:
        """
        Returns listener
        """
        return self.__listener

    # endregion Manage Events
