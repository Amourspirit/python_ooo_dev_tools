from __future__ import annotations
from typing import Any, cast, Callable, TYPE_CHECKING
import uno
from com.sun.star.frame import XStorable

from ooodev.events.calc_named_event import CalcNamedEvent
from ooodev.calc.calc_cell_range import CalcCellRange
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.utils import file_io as mFile
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.type_var import PathOrStr  # , EventCallback

from ooodev.calc.export.export_base import ExportBase

if TYPE_CHECKING:
    from ooodev.calc.filter.export_png import ExportPngT
else:
    ExportJpgT = Any


class RangePng(LoInstPropsPartial, ExportBase, EventsPartial):
    """Class for exporting cell range as a png image."""

    def __init__(self, cell_range: CalcCellRange, lo_inst: LoInst | None = None):
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        ExportBase.__init__(self)
        EventsPartial.__init__(self)
        self._cell_range = cell_range
        self._filter_name = "calc_png_Export"

    def export(self, fnm: PathOrStr, resolution: int = 96) -> None:
        """
        Exports page as jpg image.

        Args:
            fnm (PathOrStr, optional): Image file name.
            resolution (int, optional): Resolution in dpi. Defaults to 96.

        Raises:
            ValueError: If ``fnm`` is empty.
            CancelEventError: If ``EXPORTING_RANGE_PNG`` event is canceled.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~ooodev.events.calc_named_event.CalcNamedEvent.EXPORTING_RANGE_PNG` :eventref:`src-docs-event-cancel-export`
                - :py:attr:`~ooodev.events.calc_named_event.CalcNamedEvent.EXPORTED_RANGE_PNG` :eventref:`src-docs-event-export`

        Returns:
            None:

        Note:
            On exporting event is :ref:`cancel_event_args_export`.
            On exported event is :ref:`event_args_export`.
            Args ``event_data`` is a :py:class:`~ooodev.calc.filter.export_png.ExportPngT` dictionary.
        """
        if not fnm:
            raise ValueError("fnm is required")
        # if not isinstance(self._doc, StorablePartial):
        #     raise NotImplementedError(f"StorablePartial is not implemented in: {type(self._doc).__name__}")

        sz = self._cell_range.size
        dpi_x, dpi_y = self._get_dpi_width_height(sz.width.get_value_mm100(), sz.height.get_value_mm100(), resolution)

        event_data: ExportPngT = {
            "compression": 6,
            "pixel_width": dpi_x,
            "pixel_height": dpi_y,
            "interlaced": False,
            "translucent": True,
            "logical_width": dpi_x,
            "logical_height": dpi_y,
        }

        cargs = CancelEventArgsExport(source=self, event_data=event_data, fnm=fnm, overwrite=True)
        cargs.set("image_type", "png")
        cargs.set("filter_name", self._filter_name)

        self.trigger_event(CalcNamedEvent.EXPORTING_RANGE_PNG, cargs)
        if cargs.cancel and cargs.handled is False:
            raise mEx.CancelEventError(cargs)

        make_prop = mProps.Props.make_prop_value
        filter_data = [
            make_prop(name="Interlaced", value=int(cargs.event_data["interlaced"])),
            make_prop(name="Translucent", value=int(cargs.event_data["translucent"])),
        ]

        pixel_width = cargs.event_data["pixel_width"]
        pixel_height = cargs.event_data["pixel_height"]
        if pixel_width > 0 and pixel_height > 0:
            filter_data.append(make_prop(name="PixelWidth", value=pixel_width))
            filter_data.append(make_prop(name="PixelHeight", value=pixel_height))

        # compression 1..9 default 6
        compression = cargs.event_data["compression"]
        if compression > 0 and compression < 10:
            filter_data.append(make_prop(name="Compression", value=compression))

        logical_height = cargs.event_data["logical_height"]
        logical_width = cargs.event_data["logical_width"]
        if logical_height > 0 and logical_width > 0:
            filter_data.append(make_prop(name="LogicalHeight", value=logical_height))
            filter_data.append(make_prop(name="LogicalWidth", value=logical_width))

        url = mFile.FileIO.fnm_to_url(fnm=cargs.fnm)

        args = mProps.Props.make_props(
            FilterName=self._filter_name,
            FilterData=uno.Any("[]com.sun.star.beans.PropertyValue", tuple(filter_data)),  # type: ignore
            Overwrite=cargs.overwrite,
            SelectionOnly=True,
        )

        self._export_as_img(url, args)

        eargs = EventArgsExport.from_args(cargs)
        eargs.set("url", url)
        self.trigger_event(CalcNamedEvent.EXPORTED_RANGE_PNG, eargs)

    def _export_as_img(self, url: str, args: tuple) -> None:
        # capture the current selection.
        # current_sel = self._cell_range.calc_sheet.get_selected()
        # clear any selection
        self._cell_range.calc_sheet.deselect_cells()
        self._cell_range.select()

        storable = self._cell_range.calc_sheet.calc_doc.qi(XStorable, True)
        storable.storeToURL(url, args)  # save PNG
        # deselect all cells.
        self._cell_range.calc_sheet.deselect_cells()
        # restore previous selection.
        # current_sel.select()

    # region Events
    def subscribe_event_exporting(self, callback: Callable[[Any, CancelEventArgsExport[ExportPngT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on exporting.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.subscribe_event(CalcNamedEvent.EXPORTING_RANGE_PNG, callback)

    def subscribe_event_exported(self, callback: Callable[[Any, EventArgsExport[ExportPngT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on export complete.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.subscribe_event(CalcNamedEvent.EXPORTED_RANGE_PNG, callback)

    def unsubscribe_event_exporting(self, callback: Callable[[Any, CancelEventArgsExport[ExportPngT]], None]) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.unsubscribe_event(CalcNamedEvent.EXPORTING_RANGE_PNG, callback)

    def unsubscribe_event_exported(self, callback: Callable[[Any, EventArgsExport[ExportPngT]], None]) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.unsubscribe_event(CalcNamedEvent.EXPORTED_RANGE_PNG, callback)

    # endregion Events
