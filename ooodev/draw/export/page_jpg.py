from __future__ import annotations
from typing import Any, cast, Callable, TYPE_CHECKING
import uno

from ooodev.adapter.drawing.graphic_export_filter_implement import GraphicExportFilterImplement
from ooodev.events.draw_named_event import DrawNamedEvent
from ooodev.draw.draw_page import DrawPage
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import file_io as mFile
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.type_var import PathOrStr  # , EventCallback
from ooodev.draw.export.export_base import ExportBase

if TYPE_CHECKING:
    from ooodev.draw.filter.export_jpg import ExportJpgT
    from com.sun.star.lang import XComponent
else:
    ExportJpgT = Any


class PageJpg(LoInstPropsPartial, ExportBase, EventsPartial):
    """Class for exporting current Draw page as a jpg image."""

    def __init__(self, owner: DrawPage[ComponentT]):
        ExportBase.__init__(self)
        EventsPartial.__init__(self)
        self._owner = owner
        self._doc = owner.owner
        self._filter_name = "draw_jpg_Export"
        if isinstance(self._owner, LoInstPropsPartial):
            lo_inst = self._owner.lo_inst
        elif isinstance(self._doc, LoInstPropsPartial):
            lo_inst = self._doc.lo_inst
        else:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)

    # def _get_shapes(self) -> ShapeCollection:
    #     comp = ShapeCollection(self)
    #     for shape in self._draw_page:
    #         comp.add(shape.component)
    #     return comp

    def export(self, fnm: PathOrStr, resolution: int = 96) -> None:
        """
        Exports page as jpg image.

        Args:
            fnm (PathOrStr, optional): Image file name.
            resolution (int, optional): Resolution in dpi. Defaults to 96.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~ooodev.events.draw_named_event.DrawNamedEvent.EXPORTING_PAGE_JPG` :eventref:`src-docs-event-cancel-export`
                - :py:attr:`~ooodev.events.draw_named_event.DrawNamedEvent.EXPORTED_PAGE_JPG` :eventref:`src-docs-event-export`

        Returns:
            None:

        Note:
            On exporting event is :ref:`cancel_event_args_export`.
            On exported event is :ref:`event_args_export`.
            Args ``event_data`` is a :py:class:`~ooodev.draw.filter.export_jpg.ExportJpgT` dictionary.

            Unlike exporting png, exporting jpg does not seem to have a limit on the image size.

            Page margins are ignored. Any shape that is outside the page margins will not be included in the image.
        """
        if not fnm:
            raise ValueError("fnm is required")
        # if not isinstance(self._doc, StorablePartial):
        #     raise NotImplementedError(f"StorablePartial is not implemented in: {type(self._doc).__name__}")

        width = self._owner.component.Width
        height = self._owner.component.Height
        dpi_x, dpi_y = self._get_dpi_width_height(width, height, resolution)

        event_data: ExportJpgT = {
            "color_mode": True,
            "quality": 75,
            "pixel_width": dpi_x,
            "pixel_height": dpi_y,
            "logical_width": dpi_x,
            "logical_height": dpi_y,
        }

        cargs = CancelEventArgsExport(source=self, event_data=event_data, fnm=fnm, overwrite=True)
        cargs.set("image_type", "jpg")
        cargs.set("filter_name", self._filter_name)

        self.trigger_event(DrawNamedEvent.EXPORTING_PAGE_JPG, cargs)
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

        url = mFile.FileIO.fnm_to_url(fnm=cargs.fnm)

        args = mProps.Props.make_props(
            FilterName=self._filter_name,
            MediaType="image/jpeg",
            URL=url,
            FilterData=uno.Any("[]com.sun.star.beans.PropertyValue", tuple(filter_data)),  # type: ignore
            Overwrite=cargs.overwrite,
        )
        graphic_filter = GraphicExportFilterImplement(lo_inst=self.lo_inst)
        graphic_filter.set_source_document(cast("XComponent", self._owner.component))
        graphic_filter.filter(*args)

        eargs = EventArgsExport.from_args(cargs)
        eargs.set("url", url)
        self.trigger_event(DrawNamedEvent.EXPORTED_PAGE_JPG, eargs)

    # region Events
    def subscribe_event_exporting(self, callback: Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on exporting.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.subscribe_event(DrawNamedEvent.EXPORTING_PAGE_JPG, callback)

    def subscribe_event_exported(self, callback: Callable[[Any, EventArgsExport[ExportJpgT]], None]) -> None:
        """
        Add an event listener to current instance that is triggered on export complete.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.subscribe_event(DrawNamedEvent.EXPORTED_PAGE_JPG, callback)

    def unsubscribe_event_exporting(self, callback: Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, CancelEventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.unsubscribe_event(DrawNamedEvent.EXPORTING_PAGE_JPG, callback)

    def unsubscribe_event_exported(self, callback: Callable[[Any, EventArgsExport[ExportJpgT]], None]) -> None:
        """
        Remove an event listener from current instance.

        Args:
            callback (Callable[[Any, EventArgsExport[ExportJpgT]], None]): Callback of the event listener.

        Returns:
            None:
        """
        self.unsubscribe_event(DrawNamedEvent.EXPORTED_PAGE_JPG, callback)

    # endregion Events
