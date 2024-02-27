from __future__ import annotations
from typing import Any, Callable, TYPE_CHECKING
import uno

from ooodev.events.draw_named_event import DrawNamedEvent
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.draw.export.shape_export_base import ShapeExportBase

if TYPE_CHECKING:
    from ooodev.draw.filter.export_jpg import ExportJpgT
else:
    ExportJpgT = Any


class ShapeExportJpgBase(ShapeExportBase):
    """Base class for exporting shapes as a jpg image."""

    def __init__(self, lo_inst: LoInst | None = None):
        """
        Constructor

        Args:
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        ShapeExportBase.__init__(self, lo_inst=lo_inst)

    def _trigger_event_exporting(self, event_args: Any) -> None:
        self.trigger_event(DrawNamedEvent.EXPORTING_SHAPE_JPG, event_args)

    def _trigger_event_exported(self, event_args: Any) -> None:
        self.trigger_event(DrawNamedEvent.EXPORTED_SHAPE_JPG, event_args)

    # region Events
    def subscribe_event_exporting(self, callback: Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on exporting.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.subscribe_event(DrawNamedEvent.EXPORTING_SHAPE_JPG, callback)

    def subscribe_event_exported(self, callback: Callable[[Any, EventArgsExport[ExportJpgT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on export complete.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.subscribe_event(DrawNamedEvent.EXPORTED_SHAPE_JPG, callback)

    def unsubscribe_event_exporting(self, callback: Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.unsubscribe_event(DrawNamedEvent.EXPORTING_SHAPE_JPG, callback)

    def unsubscribe_event_exported(self, callback: Callable[[Any, EventArgsExport[ExportJpgT]], None]) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.unsubscribe_event(DrawNamedEvent.EXPORTED_SHAPE_JPG, callback)

    # endregion Events
