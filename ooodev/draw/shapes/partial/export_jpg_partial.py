from __future__ import annotations
from typing import Any, Callable, TYPE_CHECKING

from ooodev.draw import DrawNamedEvent
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.type_var import PathOrStr
from ooodev.events.lo_events import Events

if TYPE_CHECKING:
    from ooodev.draw.filter.export_jpg import ExportJpgT
    from com.sun.star.drawing import XShape
else:
    ExportJpgT = Any


class ExportJpgPartial:
    """Partial Class for Shapes that implements jpg export."""

    def __init__(self, component: XShape, events: Events | None = None):
        """
        Constructor.

        Args:
            component (XShape): Shape component.
            events (Events, optional): Events instance.
        """
        self.__component = component
        if events is None:
            self.__events = Events(source=self)
        else:
            self.__events = events

    def export_shape_jpg(self, fnm: PathOrStr, resolution: int = 96) -> None:
        """
        Exports shape as jpg image.

        Args:
            fnm (PathOrStr, optional): Image file name.
            resolution (int, optional): Resolution in dpi. Defaults to 96.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~ooodev.events.draw_named_event.DrawNamedEvent.EXPORTING_SHAPE_JPG` :eventref:`src-docs-event-cancel-export`
                - :py:attr:`~ooodev.events.draw_named_event.DrawNamedEvent.EXPORTED_SHAPE_JPG` :eventref:`src-docs-event-export`

        Returns:
            None:

        Note:
            On exporting event is :ref:`cancel_event_args_export`.
            On exported event is :ref:`event_args_export`.
            Args ``event_data`` is a :py:class:`~ooodev.draw.filter.export_jpg.ExportJpgT` dictionary.

            If ``fnm`` is not specified, the image file name is created based on the document name and page number
            and written to the same folder as the document.
        """
        from ooodev.draw.export.shape_jpg import ShapeJpg

        def on_exporting(source: Any, args: CancelEventArgsExport[ExportJpgT]) -> None:
            self.__events.trigger(DrawNamedEvent.EXPORTING_SHAPE_JPG, args)

        def on_exported(source: Any, args: EventArgsExport[ExportJpgT]) -> None:
            self.__events.trigger(DrawNamedEvent.EXPORTED_SHAPE_JPG, args)

        exporter = ShapeJpg(self.__component)
        exporter.subscribe_event_exporting(on_exporting)
        exporter.subscribe_event_exported(on_exported)
        exporter.export(fnm, resolution)

    # region Events
    def subscribe_event_shape_jpg_exporting(
        self, callback: Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]
    ) -> None:
        """
        Add an event listener to current instance that is triggered on exporting.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.__events.on(DrawNamedEvent.EXPORTING_SHAPE_JPG, callback)

    def subscribe_event_shape_jpg_exported(self, callback: Callable[[Any, EventArgsExport[ExportJpgT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on export complete.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.__events.on(DrawNamedEvent.EXPORTED_SHAPE_JPG, callback)

    def unsubscribe_event_shape_jpg_exporting(
        self, callback: Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]
    ) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.__events.remove(DrawNamedEvent.EXPORTING_SHAPE_JPG, callback)

    def unsubscribe_event_shape_jpg_exported(
        self, callback: Callable[[Any, EventArgsExport[ExportJpgT]], None]
    ) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.__events.remove(DrawNamedEvent.EXPORTED_SHAPE_JPG, callback)

    # endregion Events
