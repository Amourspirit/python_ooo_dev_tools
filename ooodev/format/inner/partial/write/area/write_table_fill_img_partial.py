from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
from ooodev.format.inner.style_factory import write_area_img_factory
from ooodev.utils.data_type.offset import Offset
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from com.sun.star.awt import XBitmap
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.write.fill.area.fill_img_t import FillImgT
    from ooodev.format.inner.preset.preset_image import PresetImageKind
    from ooodev.format.inner.common.format_types.size_percent import SizePercent
    from ooodev.utils.data_type.size_mm import SizeMM
    from ooo.dyn.drawing.rectangle_point import RectanglePoint
    from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
    from ooodev.format.inner.common.format_types.offset_row import OffsetRow
else:
    FillImgT = Any
    XBitmap = Any
    PresetImageKind = Any
    SizePercent = Any
    SizeMM = Any
    RectanglePoint = Any
    OffsetColumn = Any
    OffsetRow = Any

# same as write_fill_img_partial.WriteFillImgPartial except no get method


class WriteTableFillImgPartial:
    """
    Partial class for Write Fill Image.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_area_img",
            after_event="after_style_area_img",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

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
    ) -> FillImgT | None:
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
            FillImgT | None: Fill Image instance or ``None`` if cancelled.

        Hint:
            - ``RectanglePoint`` can be imported from ``ooo.dyn.drawing.rectangle_point``
            - ``ImgStyleKind`` can be imported from ``ooodev.format.inner.direct.write.fill.area.img``
            - ``SizePercent`` can be imported from ``ooodev.format.inner.common.format_types.size_percent``
            - ``SizeMM`` can be imported from ``ooodev.utils.data_type.size_mm``
            - ``OffsetColumn`` can be imported from ``ooodev.format.inner.common.format_types.offset_column``
            - ``OffsetRow`` can be imported from ``ooodev.format.inner.common.format_types.offset_row``
            - ``Offset`` can be imported from ``ooodev.utils.data_type.offset``
        """
        factory = write_area_img_factory
        kwargs = {"mode": mode, "auto_name": auto_name}
        if bitmap is not None:
            kwargs["bitmap"] = bitmap
        if name:
            kwargs["name"] = name
        if size is not None:
            kwargs["size"] = size
        if position is not None:
            kwargs["position"] = position
        if pos_offset is not None:
            kwargs["pos_offset"] = pos_offset
        if tile_offset is not None:
            kwargs["tile_offset"] = tile_offset

        return self.__styler.style(factory=factory, **kwargs)

    def style_area_image_from_preset(self, preset: PresetImageKind) -> FillImgT | None:
        """
        Style Area Gradient from Preset.

        Args:
            preset (PresetImageKind): Preset Image Kind.

        Returns:
            FillImgT: Chart Fill Image instance.

        Hint:
            - ``PresetImageKind`` can be imported from ``ooodev.format.inner.preset.preset_image``
        """
        fe = self.__styler.style_get(
            factory=write_area_img_factory,
            call_method_name="from_preset",
            event_name_suffix="_from_preset",
            obj_arg_name="",
            preset=preset,
        )
        self.__styler.style_apply(style=fe, preset=preset)
        return fe
