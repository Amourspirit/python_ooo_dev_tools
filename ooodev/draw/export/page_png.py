from __future__ import annotations
from typing import Any, Callable, TYPE_CHECKING
from pathlib import Path
import uno
from com.sun.star.frame import XStorable


from ooodev.units import UnitInch
from ooodev.draw import DrawPage
from ooodev.utils.type_var import PathOrStr  # , EventCallback
from ooodev.utils import file_io as mFile
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.draw import DrawNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.utils import info as mInfo
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.proto.component_proto import ComponentT
from ooodev.adapter.frame.storable_partial import StorablePartial

if TYPE_CHECKING:
    from ooodev.draw.filter.export_png import ExportPngT
else:
    ExportPngT = Any


class PagePng(EventsPartial):
    """Class for exporting current Writer page as a png image."""

    def __init__(self, owner: DrawPage[ComponentT]):
        EventsPartial.__init__(self)
        self._owner = owner
        self._doc = owner.owner

    def export(self, fnm: PathOrStr, resolution: int = 96) -> None:
        """
        Exports doc pages as png images.

        Args:
            fnm (PathOrStr, optional): Image file name.
            resolution (int, optional): Resolution in dpi. Defaults to 96.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.EXPORTING_PAGE_PNG` :eventref:`src-docs-event-cancel-export`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.EXPORTED_PAGE_PNG` :eventref:`src-docs-event-export`

        Returns:
            None:

        Note:
            On exporting event is :ref:`cancel_event_args_export`.
            On exported event is :ref:`event_args_export`.
            Args ``event_data`` is a :py:class:`~ooodev.write.filter.export_png.ExportPngT` dictionary.

            If ``fnm`` is not specified, the image file name is created based on the document name and page number
            and written to the same folder as the document.
        """
        # https://github.com/LibreOffice/core/blob/89e7c04ba48dab824e9f291d7db38dac6ffd6b19/svtools/source/filter/exportdialog.cxx#L783 # case FORMAT_PNG :
        # https://ask.libreoffice.org/t/export-as-png-with-macro/74337/11
        # https://ask.libreoffice.org/t/how-to-export-cell-range-to-images/57828/2
        # raises uno.com.sun.star.io.IOException if image file exists and Overwrite is False

        if not isinstance(self._doc, StorablePartial):
            raise NotImplementedError(f"StorablePartial is not implemented in: {type(self._doc).__name__}")

        width = self._owner.component.Width
        height = self._owner.component.Height
        width_in = UnitInch.from_mm100(width)
        height_in = UnitInch.from_mm100(height)
        dpi_x = round(resolution * width_in.value)
        dpi_y = round(resolution * height_in.value)

        event_data: ExportPngT = {
            "compression": 6,
            "pixel_width": dpi_x,
            "pixel_height": dpi_y,
            "interlaced": False,
            "translucent": True,
            "logical_width": width,
            "logical_height": height,
        }

        if not fnm:
            raise ValueError("fnm is required")

        cargs = CancelEventArgsExport(source=self, event_data=event_data, fnm=fnm, overwrite=True)
        cargs.set("image_type", "png")
        cargs.set("filter_name", "writer_png_Export")

        self.trigger_event(DrawNamedEvent.EXPORTING_PAGE_PNG, cargs)
        if cargs.cancel and cargs.handled is False:
            raise mEx.CancelEventError(cargs)

        make_prop = mProps.Props.make_prop_value
        filter_data = [
            make_prop(name="Interlaced", value=int(cargs.event_data["interlaced"])),
            make_prop(name="Translucent", value=int(cargs.event_data["translucent"])),
        ]
        # filter_data_props = PropertyValue("FilterData", 0, uno.Any("[]com.sun.star.beans.PropertyValue", tuple(filter_data)), dv)
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

        args = mProps.Props.make_props(
            FilterName="writer_png_Export",
            FilterData=uno.Any("[]com.sun.star.beans.PropertyValue", tuple(filter_data)),  # type: ignore
            Overwrite=cargs.overwrite,
        )
        url = mFile.FileIO.fnm_to_url(fnm=fnm)

        # - Group the shapes into a single shape.
        # - Create a new document.
        # - Copy the shape to the new document.
        # - Export the new document.
        self._doc.store_to_url(url, *args)  # save PNG

        eargs = EventArgsExport.from_args(cargs)
        eargs.set("url", url)
        self.trigger_event(DrawNamedEvent.EXPORTED_PAGE_PNG, eargs)

    # region Events
    def subscribe_event_exporting(self, callback: Callable[[Any, CancelEventArgsExport[ExportPngT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on exporting.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.subscribe_event(DrawNamedEvent.EXPORTING_PAGE_PNG, callback)

    def subscribe_event_exported(self, callback: Callable[[Any, CancelEventArgsExport[ExportPngT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on export complete.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.subscribe_event(DrawNamedEvent.EXPORTED_PAGE_PNG, callback)

    def unsubscribe_event_exporting(self, callback: Callable[[Any, CancelEventArgsExport[ExportPngT]], None]) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.unsubscribe_event(DrawNamedEvent.EXPORTING_PAGE_PNG, callback)

    def unsubscribe_event_exported(self, callback: Callable[[Any, CancelEventArgsExport[ExportPngT]], None]) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportPngT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.unsubscribe_event(DrawNamedEvent.EXPORTED_PAGE_PNG, callback)

    # endregion Events
