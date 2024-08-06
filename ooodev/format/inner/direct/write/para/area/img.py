# region Imports
from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar

from com.sun.star.awt import XBitmap

from ooo.dyn.style.graphic_location import GraphicLocation
from ooo.dyn.drawing.rectangle_point import RectanglePoint

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.events.format_named_event import FormatNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
from ooodev.format.inner.common.format_types.offset_row import OffsetRow
from ooodev.format.inner.common.format_types.size_percent import SizePercent
from ooodev.format.inner.common.props.img_para_area_props import ImgParaAreaProps
from ooodev.format.inner.direct.write.fill.area.img import Img as InnerImg
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.preset.preset_image import PresetImageKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.data_type.offset import Offset
from ooodev.utils.data_type.size_mm import SizeMM

# endregion Imports

_TImg = TypeVar("_TImg", bound="Img")


class Img(StyleMulti):
    """
    Paragraph Area Image Style

    .. seealso::

        - :ref:`help_writer_format_direct_para_area_img`

    .. versionadded:: 0.9.0
    """

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
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is required.
            name (str, optional): Specifies the name of the image. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            mode (ImgStyleKind, optional): Specifies the image style, tiled, stretched etc. Default ``ImgStyleKind.TILED``.
            size (SizePercent, SizeMM, optional): Size in percent (``0 - 100``) or size in ``mm`` units.
            position (RectanglePoint): Tiling position of Image.
            pos_offset (Offset, optional): Tiling position offset.
            tile_offset (OffsetColumn, OffsetRow, optional): The tiling offset.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.

        Returns:
            None:

        Note:
            If ``auto_name`` is ``False`` then a bitmap for a given ``name`` is only required the first call.
            All subsequent call of the same ``name`` will retrieve the bitmap form the LibreOffice Bitmap Table.

        See Also:

            - :ref:`help_writer_format_direct_para_area_img`
        """

        img = InnerImg(
            bitmap=bitmap,
            name=name,
            mode=mode,
            size=size,
            position=position,
            pos_offset=pos_offset,
            tile_offset=tile_offset,
            auto_name=auto_name,
        )
        img._prop_parent = self

        init_vars = {
            self._props.back_color: -1,
            self._props.graphic_loc: self._get_graphic_loc(position, mode),
            self._props.transparent: True,
        }
        b_map = img._get(img._props.bitmap)
        if b_map is not None:
            init_vars[self._props.back_graphic] = b_map

        super().__init__(**init_vars)
        self._set_style("fill_image", img, *img.get_attrs())  # type: ignore
        img.add_event_listener(FormatNamedEvent.STYLE_PROPERTY_APPLYING, _on_fill_img_prop_setting)
        img.add_event_listener(FormatNamedEvent.STYLE_BACKING_UP, _on_fill_img_prop_backup)

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.text.TextContent",
                "com.sun.star.beans.PropertySet",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print("FillPattern.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.FillProperties`` or ``com.sun.star.beans.PropertySet`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    # endregion apply()

    def on_property_restore_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        # mLo.Lo.print(f'Restoring "{event_args.key}"')
        defaults = (self._props.back_color, self._props.graphic_loc, self._props.transparent)
        if event_args.key in defaults:
            default = mProps.Props.get_default(event_args.event_data, event_args.key)
            if default == event_args.value:
                event_args.default = True
        elif event_args.key == self._props.back_graphic:
            if not event_args.value:
                event_args.cancel = True

        return super().on_property_restore_setting(source, event_args)

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        # mLo.Lo.print(f'Setting "{event_args.key}"')
        return super().on_property_setting(source, event_args)

    # endregion Overrides

    # region Internal methods
    def _get_graphic_loc(self, position: RectanglePoint | None, mode: ImgStyleKind | None) -> GraphicLocation:
        loc = None
        if position is not None:
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
        elif mode is not None:
            if mode == ImgStyleKind.CUSTOM:
                loc = GraphicLocation.MIDDLE_MIDDLE
            elif mode == ImgStyleKind.STRETCHED:
                loc = GraphicLocation.AREA
        if loc is None:
            loc = GraphicLocation.TILED
        return loc

    # endregion Internal methods

    # region Static Methods
    # region from_preset()
    @overload
    @classmethod
    def from_preset(cls: Type[_TImg], preset: PresetImageKind) -> _TImg: ...

    @overload
    @classmethod
    def from_preset(cls: Type[_TImg], preset: PresetImageKind, **kwargs) -> _TImg: ...

    @classmethod
    def from_preset(cls: Type[_TImg], preset: PresetImageKind, **kwargs) -> _TImg:
        """
        Gets an instance from a preset.

        Args:
            preset (PresetImageKind): Preset.

        Returns:
            Img: Instance from preset.
        """
        # pylint: disable=protected-access
        fill_img = InnerImg.from_preset(preset)

        inst = cls(**kwargs)
        inst._set(
            inst._props.graphic_loc, inst._get_graphic_loc(position=fill_img.prop_position, mode=fill_img.prop_mode)
        )
        inst._set(inst._props.back_graphic, fill_img._get(fill_img._props.bitmap))
        fill_img._prop_parent = inst
        inst._set_style("fill_image", fill_img, *fill_img.get_attrs())
        fill_img.add_event_listener(FormatNamedEvent.STYLE_PROPERTY_APPLYING, _on_fill_img_prop_setting)
        fill_img.add_event_listener(FormatNamedEvent.STYLE_BACKING_UP, _on_fill_img_prop_backup)
        return inst

    # endregion from_preset()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TImg], obj: Any) -> _TImg: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TImg], obj: Any, **kwargs) -> _TImg: ...

    @classmethod
    def from_obj(cls: Type[_TImg], obj: Any, **kwargs) -> _TImg:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Img: ``Img`` instance that represents ``obj`` fill image.
        """
        # pylint: disable=protected-access
        fill_img = InnerImg.from_obj(obj)
        inst = cls(**kwargs)
        bitmap = fill_img._get(fill_img._props.bitmap)
        inst._set(
            inst._props.graphic_loc, inst._get_graphic_loc(position=fill_img.prop_position, mode=fill_img.prop_mode)
        )
        if bitmap is not None:
            inst._set(inst._props.back_graphic, fill_img._get(fill_img._props.bitmap))
        fill_img._prop_parent = inst
        inst._set_style("fill_image", fill_img, *fill_img.get_attrs())
        fill_img.add_event_listener(FormatNamedEvent.STYLE_PROPERTY_APPLYING, _on_fill_img_prop_setting)
        fill_img.add_event_listener(FormatNamedEvent.STYLE_BACKING_UP, _on_fill_img_prop_backup)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion Static Methods
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.TXT_CONTENT | FormatKind.PARA | FormatKind.FILL
        return self._format_kind_prop

    @property
    def prop_inner(self) -> InnerImg:
        """Gets Fill image instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerImg, self._get_style_inst("fill_image"))
        return self._direct_inner

    @property
    def _props(self) -> ImgParaAreaProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = ImgParaAreaProps(
                back_color="ParaBackColor",
                back_graphic="ParaBackGraphic",
                graphic_loc="ParaBackGraphicLocation",
                transparent="ParaBackTransparent",
            )
        return self._props_internal_attributes


def _on_fill_img_prop_setting(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    # pylint: disable=protected-access
    # Fill images do not support setting of FillBitmapMode in Write
    img = cast(InnerImg, event_args.event_source)
    if event_args.key == img._props.mode:
        # print("Setting Fill Property: Found FillBitmapMode, canceling.")
        event_args.cancel = True


def _on_fill_img_prop_backup(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    # Fill images do not support setting of FillBitmapMode in Write
    img = cast(InnerImg, event_args.event_source)
    if event_args.key == img._props.mode:
        # print("Backup up Fill Property: Found FillBitmapMode, canceling.")
        event_args.cancel = True
