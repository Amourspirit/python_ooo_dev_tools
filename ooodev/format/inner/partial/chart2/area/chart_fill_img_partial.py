from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
from ooodev.format.inner.style_factory import chart2_area_img_factory
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.data_type.offset import Offset

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from com.sun.star.awt import XBitmap
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.chart2.area.chart_fill_img_t import ChartFillImgT
    from ooodev.format.inner.preset.preset_image import PresetImageKind
    from ooodev.format.inner.common.format_types.size_percent import SizePercent
    from ooodev.utils.data_type.size_mm import SizeMM
    from ooo.dyn.drawing.rectangle_point import RectanglePoint
    from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
    from ooodev.format.inner.common.format_types.offset_row import OffsetRow
else:
    XChartDocument = Any
    XBitmap = Any
    PresetImageKind = Any
    SizePercent = Any
    SizeMM = Any
    RectanglePoint = Any
    OffsetColumn = Any
    OffsetRow = Any


class ChartFillImgPartial:
    """
    Partial class for Chart Fill Image.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

    def _ChartFillImgPartial__get_chart_doc(self) -> XChartDocument:
        if isinstance(self, ChartDocPropPartial):
            return self.chart_doc.component

        raise NotImplementedError

    def style_area_image(
        self,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        mode: ImgStyleKind = ImgStyleKind.TILED,
        size: SizePercent | SizeMM | None = None,
        position: RectanglePoint | None = None,
        pos_offset: Offset | None = None,
        tile_offset: OffsetColumn | OffsetRow | None = None,
        auto_name: bool = False,
    ) -> ChartFillImgT | None:
        """
        Style Area Color.

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table
                then this property is required.
            name (str, optional): Specifies the name of the image. This is also the name that is used to store
                bitmap in LibreOffice Bitmap Table.
            mode (~.write.fill.area.img.ImgStyleKind, optional): Specifies the image style, tiled, stretched etc.
                Default ``ImgStyleKind.TILED``.
            size (SizePercent, SizeMM, optional): Size in percent (``0 - 100``) or size in ``mm`` units.
            position (RectanglePoint): Tiling position of Image.
            pos_offset (Offset, optional): Tiling position offset.
            tile_offset (OffsetColumn, OffsetRow, optional): The tiling offset.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.

        Raises:
            CancelEventError: If the event ``before_style_area_img`` is cancelled and not handled.

        Returns:
            ChartFillImgT | None: Fill Image instance or ``None`` if cancelled.

        Hint:
            - ``RectanglePoint`` can be imported from ``ooo.dyn.drawing.rectangle_point``
            - ``ImgStyleKind`` can be imported from ``ooodev.format.inner.direct.write.fill.area.img``
            - ``SizePercent`` can be imported from ``ooodev.format.inner.common.format_types.size_percent``
            - ``SizeMM`` can be imported from ``ooodev.utils.data_type.size_mm``
            - ``OffsetColumn`` can be imported from ``ooodev.format.inner.common.format_types.offset_column``
            - ``OffsetRow`` can be imported from ``ooodev.format.inner.common.format_types.offset_row``
            - ``Offset`` can be imported from ``ooodev.utils.data_type.offset``
        """
        doc = self._ChartFillImgPartial__get_chart_doc()
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_area_image.__qualname__)
            event_data: Dict[str, Any] = {
                "bitmap": bitmap,
                "name": name,
                "mode": mode,
                "size": size,
                "position": position,
                "pos_offset": pos_offset,
                "tile_offset": tile_offset,
                "auto_name": auto_name,
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_area_img", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_area_img")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Area Image has been cancelled.")
                    else:
                        return None
                else:
                    return None
            bitmap = cargs.event_data.get("bitmap", bitmap)
            name = cargs.event_data.get("name", name)
            mode = cargs.event_data.get("mode", mode)
            size = cargs.event_data.get("size", size)
            position = cargs.event_data.get("position", position)
            pos_offset = cargs.event_data.get("pos_offset", pos_offset)
            tile_offset = cargs.event_data.get("tile_offset", tile_offset)
            auto_name = cargs.event_data.get("auto_name", auto_name)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = chart2_area_img_factory(factory_name)
        fe = styler(
            chart_doc=doc,
            bitmap=bitmap,
            name=name,
            mode=mode,
            size=size,
            position=position,
            pos_offset=pos_offset,
            tile_offset=tile_offset,
            auto_name=auto_name,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_area_img", EventArgs.from_args(cargs))  # type: ignore
        return fe

    def style_area_image_get(self) -> ChartFillImgT | None:
        """
        Gets the Area Area Image Style.

        Raises:
            CancelEventError: If the event ``before_style_area_img_get`` is cancelled and not handled.

        Returns:
            ChartFillImgT | None: Area image style or ``None`` if cancelled.
        """
        doc = self._ChartFillImgPartial__get_chart_doc()
        comp = self.__component
        factory_name = self.__factory_name
        cargs = None
        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self.style_area_image_get.__qualname__)
            event_data: Dict[str, Any] = {
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_area_img_get", cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", "before_style_area_img_get")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style get has been cancelled.")
                else:
                    return None
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = chart2_area_img_factory(factory_name)
        try:
            style = styler.from_obj(chart_doc=doc, obj=comp)
        except mEx.DisabledMethodError:
            return None

        style.set_update_obj(comp)
        return style

    def style_area_image_from_preset(self, preset: PresetImageKind) -> ChartFillImgT | None:
        """
        Style Area Gradient from Preset.

        Args:
            preset (PresetImageKind): Preset Image Kind.

        Returns:
            ChartFillImgT: Chart Fill Image instance.

        Hint:
            - ``PresetImageKind`` can be imported from ``ooodev.format.inner.preset.preset_image``
        """
        doc = self._ChartFillImgPartial__get_chart_doc()
        cargs = None
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        comp = self.__component
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_area_image_from_preset.__qualname__)
            event_data: Dict[str, Any] = {
                "preset": preset,
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_area_img_from_preset", cargs)
            if cargs.cancel is True:
                if cargs.handled is not False:
                    return None
                cargs.set("initial_event", "before_style_area_img_from_preset")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style Area Image has been cancelled.")
                else:
                    return None
            preset = cargs.event_data.get("preset", preset)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = chart2_area_img_factory(factory_name)
        fe = styler.from_preset(chart_doc=doc, preset=preset)

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_area_img_from_preset", EventArgs.from_args(cargs))  # type: ignore
        return fe
