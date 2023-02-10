from __future__ import annotations
from typing import Any, Tuple, overload

import uno

from .....events.args.key_val_cancel_args import KeyValCancelArgs
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from ....style_base import StyleMulti
from .....utils.data_type.offset import Offset as Offset
from ...fill.area.img import (
    Img as FillImg,
    SizeMM as SizeMM,
    SizePercent as SizePercent,
    OffsetColumn as OffsetColumn,
    OffsetRow as OffsetRow,
    ImgStyleKind as ImgStyleKind,
)
from ....preset.preset_image import PresetImageKind as PresetImageKind
from ....kind.format_kind import FormatKind
from .....events.format_named_event import FormatNamedEvent

from com.sun.star.awt import XBitmap

from ooo.dyn.style.graphic_location import GraphicLocation
from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint


class Img(StyleMulti):
    def __init__(
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
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is requied.
            name (str, optional): Specifies the name of the image. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            mode (ImgStyleKind, optional): Specifies the image style, tiled, stretched etc. Default ``ImgStyleKind.TILED``.
            size (SizePercent, SizeMM, optional): Size in percent (``0 - 100``) or size in ``mm`` units.
            positin (RectanglePoint): Tiling position of Image.
            pos_offset (Offset, optional): Tiling position offset.
            tile_offset (OffsetColumn, OffsetRow, optional): The tiling offset.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.

        Returns:
            None:

        Note:
            If ``auto_name`` is ``False`` then a bitmap for a given ``name`` is only required the first call.
            All subsequent call of the same ``name`` will retreive the bitmap form the LibreOffice Bitmap Table.
        """

        fimg = FillImg(
            bitmap=bitmap,
            name=name,
            mode=mode,
            size=size,
            position=position,
            pos_offset=pos_offset,
            tile_offset=tile_offset,
            auto_name=auto_name,
        )

        init_vars = {
            "ParaBackColor": -1,
            "ParaBackGraphicLocation": self._get_graphic_loc(position, mode),
            "ParaBackTransparent": True,
        }
        bmap = fimg._get("FillBitmap")
        if not bmap is None:
            init_vars["ParaBackGraphic"] = bmap

        super().__init__(**init_vars)
        self._set_style("fill_image", fimg, *fimg.get_attrs())
        fimg.add_event_listener(FormatNamedEvent.STYLE_PROPERTY_APPLYING, _on_fill_img_prop_setting)
        fimg.add_event_listener(FormatNamedEvent.STYLE_BACKING_UP, _on_fill_img_prop_backup)

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.text.TextContent",
            "com.sun.star.beans.PropertySet",
            "com.sun.star.style.ParagraphStyle",
        )

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"FillPattern.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.FillProperties`` or ``com.sun.star.beans.PropertySet`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"Img.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    def on_property_restore_setting(self, event_args: KeyValCancelArgs) -> None:
        # mLo.Lo.print(f'Restoring "{event_args.key}"')
        defaults = ("ParaBackColor", "ParaBackGraphicLocation", "ParaBackTransparent")
        if event_args.key in defaults:
            default = mProps.Props.get_default(event_args.event_data, event_args.key)
            if default == event_args.value:
                event_args.default = True
        elif event_args.key == "ParaBackGraphic":
            if not event_args.value:
                event_args.cancel = True

        return super().on_property_restore_setting(event_args)

    def on_property_setting(self, event_args: KeyValCancelArgs) -> None:
        # mLo.Lo.print(f'Setting "{event_args.key}"')
        return super().on_property_setting(event_args)

    # endregion Overrides

    # region Internal methods
    def _get_graphic_loc(self, position: RectanglePoint | None, mode: ImgStyleKind | None) -> GraphicLocation:
        loc = None
        if not position is None:
            if position == RectanglePoint.LEFT_BOTTOM:
                loc = GraphicLocation.LEFT_BOTTOM
            elif position == RectanglePoint.LEFT_MIDDLE:
                loc = GraphicLocation.LEFT_MIDDLE
            elif position == RectanglePoint.LEFT_TOP:
                loc = GraphicLocation.LEFT_TOP
            elif position == RectanglePoint.MIDDLE_BOTTOM:
                loc = GraphicLocation.MIDDLE_BOTTOM
            elif position == RectanglePoint.MIDDLE_MIDDLE:
                loc = GraphicLocation.MIDDLE_MIDDLE
            elif position == RectanglePoint.MIDDLE_TOP:
                loc = GraphicLocation.MIDDLE_TOP
            elif position == RectanglePoint.RIGHT_BOTTOM:
                loc = GraphicLocation.RIGHT_BOTTOM
            elif position == RectanglePoint.RIGHT_MIDDLE:
                loc = GraphicLocation.RIGHT_MIDDLE
            elif position == RectanglePoint.RIGHT_TOP:
                loc = GraphicLocation.RIGHT_TOP
        elif not mode is None:
            if mode == ImgStyleKind.CUSTOM:
                loc = GraphicLocation.MIDDLE_MIDDLE
            elif mode == ImgStyleKind.STRETCHED:
                loc = GraphicLocation.AREA
        if loc is None:
            loc = GraphicLocation.TILED
        return loc

    # endregion Internal methods

    # region Static Methods
    @classmethod
    def from_preset(cls, preset: PresetImageKind) -> Img:
        """
        Gets an instance from a preset

        Args:
            preset (PresetImageKind): Preset

        Returns:
            Img: Instance from preset.
        """
        fill_img = FillImg.from_preset(preset)

        inst = super(Img, cls).__new__(cls)
        inst.__init__()
        inst._set(
            "ParaBackGraphicLocation", inst._get_graphic_loc(position=fill_img.prop_posiion, mode=fill_img.prop_mode)
        )
        inst._set("ParaBackGraphic", fill_img._get("FillBitmap"))
        inst._set_style("fill_image", fill_img, *fill_img.get_attrs())
        fill_img.add_event_listener(FormatNamedEvent.STYLE_PROPERTY_APPLYING, _on_fill_img_prop_setting)
        fill_img.add_event_listener(FormatNamedEvent.STYLE_BACKING_UP, _on_fill_img_prop_backup)
        return inst

    @classmethod
    def from_obj(cls, obj: object) -> Img:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Img: ``Img`` instance that represents ``obj`` fill imgage.
        """
        fill_img = FillImg.from_obj(obj)
        inst = super(Img, cls).__new__(cls)
        inst.__init__()
        bmap = fill_img._get("FillBitmap")
        inst._set(
            "ParaBackGraphicLocation", inst._get_graphic_loc(position=fill_img.prop_posiion, mode=fill_img.prop_mode)
        )
        if not bmap is None:
            inst._set("ParaBackGraphic", fill_img._get("FillBitmap"))
        inst._set_style("fill_image", fill_img, *fill_img.get_attrs())
        fill_img.add_event_listener(FormatNamedEvent.STYLE_PROPERTY_APPLYING, _on_fill_img_prop_setting)
        fill_img.add_event_listener(FormatNamedEvent.STYLE_BACKING_UP, _on_fill_img_prop_backup)
        return inst

    # endregion Static Methods

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.TXT_CONTENT | FormatKind.PARA | FormatKind.FILL


def _on_fill_img_prop_setting(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    # Fill images do not support setting of FillBitmapMode in Write
    if event_args.key == "FillBitmapMode":
        # print("Setting Fill Property: Found FillBitmapMode, canceling.")
        event_args.cancel = True


def _on_fill_img_prop_backup(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    # Fill images do not support setting of FillBitmapMode in Write
    if event_args.key == "FillBitmapMode":
        # print("Backup up Fill Property: Found FillBitmapMode, canceling.")
        event_args.cancel = True
