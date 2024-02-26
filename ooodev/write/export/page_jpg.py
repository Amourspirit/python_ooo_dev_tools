from __future__ import annotations
from typing import Any, Callable, TYPE_CHECKING
from pathlib import Path
import uno
from com.sun.star.frame import XStorable

from ooodev.units.unit_inch import UnitInch
from ooodev.write.write_doc import WriteDoc
from ooodev.utils.type_var import PathOrStr  # , EventCallback
from ooodev.utils import file_io as mFile
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.write_named_event import WriteNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from ooodev.write.filter.export_jpg import ExportJpgT
else:
    ExportJpgT = Any


class PageJpg(LoInstPropsPartial, EventsPartial):
    """Class for exporting current Writer page as a jpg image."""

    def __init__(self, doc: WriteDoc):
        self._doc = doc
        LoInstPropsPartial.__init__(self, lo_inst=self._doc.lo_inst)
        EventsPartial.__init__(self)

    def export(self, fnm: PathOrStr = "", resolution: int = 96) -> None:
        """
        Exports doc pages as jpg images.

        Args:
            fnm (PathOrStr, optional): Image file name.
            resolution (int, optional): Resolution in dpi. Defaults to 96.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.EXPORTING_PAGE_JPG` :eventref:`src-docs-event-cancel-export`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.EXPORTED_PAGE_JPG` :eventref:`src-docs-event-export`

        Returns:
            None:

        Note:
            On exporting event is :ref:`cancel_event_args_export`.
            On exported event is :ref:`event_args_export`.
            Args ``event_data`` is a :py:class:`~ooodev.write.filter.export_jpg.ExportJpgT` dictionary.

            If ``fnm`` is not specified, the image file name is created based on the document name and page number
            and written to the same folder as the document.
        """
        # raises uno.com.sun.star.io.IOException if image file exists and Overwrite is False

        sz = self._doc.get_page_size()
        width_in = UnitInch.from_mm100(sz.width)
        height_in = UnitInch.from_mm100(sz.height)
        dpi_x = round(resolution * width_in.value)
        dpi_y = round(resolution * height_in.value)

        event_data: ExportJpgT = {
            "color_mode": True,
            "quality": 75,
            "pixel_width": dpi_x,
            "pixel_height": dpi_y,
            "logical_width": sz.width,
            "logical_height": sz.height,
        }

        view = self._doc.get_view_cursor()

        if fnm == "":
            # create a file name based on the document name and page number
            doc_path = Path(uno.fileUrlToSystemPath(self._doc.component.URL))  # type: ignore
            # create a page file number and pad with 3 leading zeros
            # 001, 002, 003, etc.
            pg_num = view.get_current_page_num()
            pg_num_str = str(pg_num).zfill(3)
            fnm = doc_path.parent / f"{doc_path.stem}_page_{pg_num_str}.jpg"

        cargs = CancelEventArgsExport(source=self, event_data=event_data, fnm=fnm, overwrite=True)
        cargs.set("image_type", "jpg")
        cargs.set("filter_name", "writer_jpg_Export")

        self.trigger_event(WriteNamedEvent.EXPORTING_PAGE_JPG, cargs)
        if cargs.cancel and cargs.handled is False:
            raise mEx.CancelEventError(cargs)

        make_prop = mProps.Props.make_prop_value
        filter_data = [
            make_prop(name="ColorMode", value=int(not cargs.event_data["color_mode"])),
        ]
        # filter_data_props = PropertyValue("FilterData", 0, uno.Any("[]com.sun.star.beans.PropertyValue", tuple(filter_data)), dv)
        pixel_width = cargs.event_data["pixel_width"]
        pixel_height = cargs.event_data["pixel_height"]
        if pixel_width > 0 and pixel_height > 0:
            filter_data.append(make_prop(name="PixelWidth", value=pixel_width))
            filter_data.append(make_prop(name="PixelHeight", value=pixel_height))

        # quality 1..100 default 75
        quality = cargs.event_data["quality"]
        if quality > 0 and quality < 101:
            filter_data.append(make_prop(name="Quality", value=quality))

        logical_height = cargs.event_data["logical_height"]
        logical_width = cargs.event_data["logical_width"]
        if logical_height > 0 and logical_width > 0:
            filter_data.append(make_prop(name="LogicalHeight", value=logical_height))
            filter_data.append(make_prop(name="LogicalWidth", value=logical_width))

        args = mProps.Props.make_props(
            FilterName="writer_jpg_Export",
            FilterData=uno.Any("[]com.sun.star.beans.PropertyValue", tuple(filter_data)),  # type: ignore
            Overwrite=cargs.overwrite,
        )
        url = mFile.FileIO.fnm_to_url(fnm=fnm)

        # view.jump_to_first_page() # export current page not first page
        storable = self._doc.qi(XStorable, True)
        storable.storeToURL(url, args)  # save PNG

        eargs = EventArgsExport.from_args(cargs)
        eargs.set("url", url)
        self.trigger_event(WriteNamedEvent.EXPORTED_PAGE_JPG, eargs)

    # region Events
    def subscribe_event_exporting(self, callback: Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on exporting.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.subscribe_event(WriteNamedEvent.EXPORTING_PAGE_JPG, callback)

    def subscribe_event_exported(self, callback: Callable[[Any, EventArgsExport[ExportJpgT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on export complete.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.subscribe_event(WriteNamedEvent.EXPORTED_PAGE_JPG, callback)

    def unsubscribe_event_exporting(self, callback: Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.unsubscribe_event(WriteNamedEvent.EXPORTING_PAGE_JPG, callback)

    def unsubscribe_event_exported(self, callback: Callable[[Any, EventArgsExport[ExportJpgT]], None]) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.unsubscribe_event(WriteNamedEvent.EXPORTED_PAGE_JPG, callback)

    # endregion Events
