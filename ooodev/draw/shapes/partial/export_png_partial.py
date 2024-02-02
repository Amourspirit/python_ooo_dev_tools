from __future__ import annotations
from typing import Any, Callable, TYPE_CHECKING

from ooodev.draw import DrawNamedEvent
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.events.lo_events import Events
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.type_var import PathOrStr

if TYPE_CHECKING:
    from ooodev.draw.filter.export_png import ExportPngT
    from com.sun.star.drawing import XShape
else:
    ExportPngT = Any


class ExportPngPartial:
    """Partial Class for Shapes that implements png export."""

    def __init__(self, component: XShape, events: Events | None = None, lo_inst: LoInst | None = None):
        """
        Constructor.

        Args:
            component (XShape): Shape component.
            events (Events, optional): Events instance.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        self.__component = component
        if events is None:
            self.__events = Events(source=self)
        else:
            self.__events = events
        if lo_inst is None:
            self.__lo_inst = mLo.Lo.current_lo
        else:
            self.__lo_inst = lo_inst

    def export_shape_png(self, fnm: PathOrStr, resolution: int = 96) -> None:
        """
        Exports shape as png image.

        Args:
            fnm (PathOrStr, optional): Image file name.
            resolution (int, optional): Resolution in dpi. Defaults to 96.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~ooodev.events.draw_named_event.DrawNamedEvent.EXPORTING_SHAPE_PNG` :eventref:`src-docs-event-cancel-export`
                - :py:attr:`~ooodev.events.draw_named_event.DrawNamedEvent.EXPORTED_SHAPE_PNG` :eventref:`src-docs-event-export`

        Returns:
            None:

        Note:
            On exporting event is :ref:`cancel_event_args_export`.
            On exported event is :ref:`event_args_export`.
            Args ``event_data`` is a :py:class:`~ooodev.draw.filter.export_png.ExportPngT` dictionary.

            If ``fnm`` is not specified, the image file name is created based on the document name and page number
            and written to the same folder as the document.

            When page is exported as png, such as an impress slide, any images on the will not be exported if filter ``translucent=True``.
            For this reason, the default value for ``translucent`` is ``False``.
        """
        from ooodev.draw.export.shape_png import ShapePng

        def on_exporting(source: Any, args: CancelEventArgsExport[ExportPngT]) -> None:
            self.__events.trigger(DrawNamedEvent.EXPORTING_SHAPE_PNG, args)

        def on_exported(source: Any, args: EventArgsExport[ExportPngT]) -> None:
            self.__events.trigger(DrawNamedEvent.EXPORTED_SHAPE_PNG, args)

        exporter = ShapePng(shape=self.__component, lo_inst=self.__lo_inst)
        exporter.subscribe_event_exporting(on_exporting)
        exporter.subscribe_event_exported(on_exported)
        exporter.export(fnm, resolution)

    # region Events
    def subscribe_event_shape_png_exporting(
        self, callback: Callable[[Any, CancelEventArgsExport[ExportPngT]], None]
    ) -> None:
        """
        Add an event listener to current instance that is triggered on exporting.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.__events.on(DrawNamedEvent.EXPORTING_SHAPE_PNG, callback)

    def subscribe_event_shape_png_exported(self, callback: Callable[[Any, EventArgsExport[ExportPngT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on export complete.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.__events.on(DrawNamedEvent.EXPORTED_SHAPE_PNG, callback)

    def unsubscribe_event_shape_png_exporting(
        self, callback: Callable[[Any, CancelEventArgsExport[ExportPngT]], None]
    ) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.__events.remove(DrawNamedEvent.EXPORTING_SHAPE_PNG, callback)

    def unsubscribe_event_shape_png_exported(
        self, callback: Callable[[Any, EventArgsExport[ExportPngT]], None]
    ) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.__events.remove(DrawNamedEvent.EXPORTED_SHAPE_PNG, callback)

    # endregion Events
