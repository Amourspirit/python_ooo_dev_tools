from __future__ import annotations
from typing import Any
import uno
from com.sun.star.lang import XComponent
from com.sun.star.drawing import XShape

from ooodev.adapter.drawing.graphic_export_filter_implement import GraphicExportFilterImplement
from ooodev.draw.filter.export_jpg import ExportJpg
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.exceptions import ex as mEx
from ooodev.utils import file_io as mFile
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.type_var import PathOrStr
from ooodev.draw.export.shape_export_jpg_base import ShapeExportJpgBase


class ShapeJpg(ShapeExportJpgBase):
    """Class for exporting current Draw page as a jpg image."""

    def __init__(self, shape: Any, lo_inst: LoInst | None = None):
        """
        Constructor

        Args:
            shape (Any): Shape to export as image.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        ShapeExportJpgBase.__init__(self, lo_inst=lo_inst)
        self._component = mLo.Lo.qi(XComponent, shape, True)
        self._filter_name = "draw_jpg_Export"

    def export(self, fnm: PathOrStr, resolution: int = 96) -> None:
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
        # raises uno.com.sun.star.io.IOException if image file exists and Overwrite is False

        # Although it is not documented in the Developer's Guide, as of OOo 1.1, it is possible to
        # export an image at a specified resolution. The MediaDescriptor in each graphic filter supports
        # the FilterData property sequence, which sets the image size in pixels using the properties
        # PixelWidth and PixelHeight. The logical size can be set in units of 1/100 mm using the
        # properties LogicalWidth and LogicalHeight.

        # This uses the GraphicExportFilter, which is only able to export a shape, shapes, or a draw page.

        if not fnm:
            raise ValueError("fnm is required")

        shape = mLo.Lo.qi(XShape, self._component, True)
        sz = shape.getSize()
        pixel_width, pixel_height = self._get_dpi_width_height(sz.Width, sz.Height, resolution)

        event_data = ExportJpg(
            color_mode=True,
            quality=75,
            pixel_width=pixel_width,
            pixel_height=pixel_height,
            logical_width=pixel_width,
            logical_height=pixel_height,
        )

        cargs = CancelEventArgsExport(source=self, event_data=event_data, fnm=fnm, overwrite=True)
        cargs.set("image_type", "jpg")
        cargs.set("filter_name", self._filter_name)

        self._trigger_event_exporting(cargs)
        if cargs.cancel and cargs.handled is False:
            raise mEx.CancelEventError(cargs)

        make_prop = mProps.Props.make_prop_value
        filter_data = [
            make_prop(name="ColorMode", value=int(not cargs.event_data["color_mode"])),
        ]
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

        # filter_data.append(make_prop(name="ExportMode", value=1))
        # filter_data.append(make_prop(name="Resolution", value=resolution))

        url = mFile.FileIO.fnm_to_url(fnm=fnm)
        args = mProps.Props.make_props(
            FilterName=self._filter_name,
            MediaType="image/jpeg",
            URL=url,
            FilterData=uno.Any("[]com.sun.star.beans.PropertyValue", tuple(filter_data)),  # type: ignore
            Overwrite=cargs.overwrite,
        )
        graphic_filter = GraphicExportFilterImplement(lo_inst=self.lo_inst)
        graphic_filter.set_source_document(self._component)
        graphic_filter.filter(*args)

        eargs = EventArgsExport.from_args(cargs)
        eargs.set("url", url)
        self._trigger_event_exported(eargs)
