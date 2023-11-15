from __future__ import annotations

from .paint_listener import PaintListener
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.utils.type_var import EventArgsCallbackT, ListenerEventCallbackT


class PaintEvents:
    """
    Class for managing Paint Events.

    This class is usually inherited by control classes that implement ``com.sun.star.awt.XPaintListener``.
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
        self.__paint_listener = PaintListener(trigger_args=trigger_args)

    # region Manage Events
    def add_event_window_paint(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event Is invoked when a region of the window became invalid, e.g.
        when another window has been moved away.

        The callback ``EventArgs.event_data`` will contain a UNO ``PaintEvent`` struct.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowPaint")
            self.__callback(self, args)
        self.__paint_listener.on("windowPaint", cb)

    def remove_event_window_paint(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="windowPaint", is_add=False)
            self.__callback(self, args)
        self.__paint_listener.off("windowPaint", cb)

    # endregion Manage Events

    @property
    def events_listener_paint(self) -> PaintListener:
        """
        Returns listener
        """
        return self.__paint_listener
