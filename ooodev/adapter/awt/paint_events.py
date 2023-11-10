from __future__ import annotations

from .paint_listener import PaintListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.utils.type_var import EventArgsCallbackT


class PaintEvents:
    def __init__(self, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        self.__paint_listener = PaintListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_window_paint(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event Is invoked when a region of the window became invalid, e.g.
        when another window has been moved away.

        The callback ``EventArgs.event_data`` will contain a UNO ``PaintEvent`` struct.
        """
        self.__paint_listener.on("windowPaint", cb)

    def remove_event_window_paint(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__paint_listener.off("windowPaint", cb)

    # endregion Manage Events

    @property
    def events_listener_paint(self) -> PaintListener:
        """
        Returns listener
        """
        return self.__paint_listener
