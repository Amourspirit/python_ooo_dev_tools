"""
Module for Fill Properties Fill Image.

.. versionadded:: 0.9.0
"""

# region Imports
from __future__ import annotations
import contextlib
from typing import Any, Tuple, cast, overload, Type, TypeVar
from enum import Enum

from com.sun.star.awt import XBitmap

from ooo.dyn.drawing.bitmap_mode import BitmapMode
from ooo.dyn.drawing.fill_style import FillStyle
from ooo.dyn.drawing.rectangle_point import RectanglePoint

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
from ooodev.format.inner.common.format_types.offset_row import OffsetRow
from ooodev.format.inner.common.format_types.size_percent import SizePercent
from ooodev.format.inner.common.props.area_img_props import AreaImgProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.preset import preset_image as mImage
from ooodev.format.inner.preset.preset_image import PresetImageKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.units.unit_convert import UnitConvert
from ooodev.utils import props as mProps
from ooodev.utils.data_type.offset import Offset
from ooodev.utils.data_type.size_mm import SizeMM

# endregion Imports


# https://github.com/LibreOffice/core/blob/6379414ca34527fbe69df2035d49d651655317cd/vcl/source/filter/ipict/ipict.cxx#L92

_TImg = TypeVar("_TImg", bound="Img")


class ImgStyleKind(Enum):
    """Image Style Kind for Image Fill"""

    CUSTOM = BitmapMode.NO_REPEAT
    """Image does not repeat"""
    TILED = BitmapMode.REPEAT
    """Image Repeats"""
    STRETCHED = BitmapMode.STRETCH
    """Image is stretched"""

    @staticmethod
    def from_bitmap_mode(value: BitmapMode) -> ImgStyleKind:
        """Gets ``ImgStyleKind`` from ``BitmapMode``"""
        if value == BitmapMode.REPEAT:
            return ImgStyleKind.TILED
        if value == BitmapMode.STRETCH:
            return ImgStyleKind.STRETCHED
        return ImgStyleKind.CUSTOM


class Img(StyleBase):
    """
    Class for Fill Properties Fill Image.

    .. seealso::

        - :ref:`help_writer_format_direct_shape_image`

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
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table
                then this property is required.
            name (str, optional): Specifies the name of the image. This is also the name that is used to store
                bitmap in LibreOffice Bitmap Table.
            mode (ImgStyleKind, optional): Specifies the image style, tiled, stretched etc.
                Default ``ImgStyleKind.TILED``.
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

            - :ref:`help_writer_format_direct_shape_image`
        """

        # when mode is ImgStyleKind.STRETCHED size, position, pos_offset, and tile_offset are not required

        init_vals = {}
        b_map = None
        with contextlib.suppress(Exception):
            # if bitmap or name is passed in then get the bitmap
            b_map = self._get_bitmap(bitmap, name, auto_name)
        if b_map is not None:
            if self._props.bitmap:
                init_vals[self._props.bitmap] = b_map
            init_vals[self._props.name] = self._name
            init_vals[self._props.style] = FillStyle.BITMAP

        super().__init__(**init_vals)
        self.prop_mode = mode
        self.prop_size = size
        self.prop_position = position
        self.prop_pos_offset = pos_offset
        self.prop_tile_offset = tile_offset

    # region Internal Methods
    def _get_bitmap(self, bitmap: XBitmap | None, name: str, auto_name: bool) -> XBitmap:
        # if the name passed in already exist in the Bitmap Table then it is returned.
        # Otherwise the bitmap is added to the Bitmap Table and then returned.
        # after bitmap is added to table all other subsequent call of this name will return
        # that bitmap from the Table. With the exception of auto_name which will force a new entry
        # into the Table each time.
        if not name:
            auto_name = True
            name = "Bitmap "
        nc = self._container_get_inst()
        if auto_name:
            self._name = self._container_get_unique_el_name(name, nc)
        else:
            self._name = name
        b_map = self._container_get_value(self._name, nc)  # raises value error if name is empty
        if b_map is not None:
            return b_map
        if bitmap is None:
            raise ValueError(f'No bitmap could be found in container for "{name}". In this case a bitmap is required.')
        self._container_add_value(name=self._name, obj=bitmap, allow_update=False, nc=nc)
        return self._container_get_value(self._name, nc)

    # endregion Internal Methods

    # region Overrides

    # region copy()
    @overload
    def copy(self: _TImg) -> _TImg: ...

    @overload
    def copy(self: _TImg, **kwargs) -> _TImg: ...

    def copy(self: _TImg, **kwargs) -> _TImg:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._name = self._name
        return cp

    # endregion copy()

    def _container_get_service_name(self) -> str:
        # https://github.com/LibreOffice/core/blob/d9e044f04ac11b76b9a3dac575f4e9155b67490e/chart2/source/tools/PropertyHelper.cxx#L246
        return "com.sun.star.drawing.BitmapTable"

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.beans.PropertySet",
                "com.sun.star.drawing.FillProperties",
                "com.sun.star.style.PageStyle",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextContent",
                "com.sun.star.text.TextEmbeddedObject",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
            )
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.FillProperties``
                or ``com.sun.star.beans.PropertySet`` service.

        Returns:
            None:
        """

        if self._props.bitmap and not self._has(self._props.bitmap):
            mLo.Lo.print("Img.apply(): There is nothing to apply.")
            return
        super().apply(obj, **kwargs)

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print("Img.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    def on_property_restore_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        if self._props.bitmap and event_args.key == self._props.bitmap:
            if event_args.value is None:
                event_args.default = True
        elif event_args.key == self._props.name:
            if not event_args.value:
                event_args.default = True
        return super().on_property_restore_setting(source, event_args)

    # endregion Overrides

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
            preset (~.preset.preset_image.PresetImageKind): Preset.

        Returns:
            Img: Instance from preset.
        """
        name = str(preset)
        nu = cls(**kwargs)

        nc = nu._container_get_inst()
        bitmap = cast(XBitmap, nu._container_get_value(name, nc))
        if bitmap is None:
            bitmap = mImage.get_prest_bitmap(preset)
        inst = cls(
            bitmap=bitmap,
            name=name,
            mode=ImgStyleKind.TILED,
            position=RectanglePoint.MIDDLE_MIDDLE,
            pos_offset=Offset(0, 0),
            tile_offset=OffsetRow(0),
            auto_name=False,
            **kwargs,
        )
        # set size
        point = preset._get_point()
        inst._set(inst._props.size_x, point.x)
        inst._set(inst._props.size_y, point.y)
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
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not support to convert to Img")

        def set_prop(key: str, fp: Img):
            nonlocal obj
            if not key:
                return
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                fp._set(key, val)

        name = mProps.Props.get(obj, inst._props.name)
        inst._name = name
        inst._set(inst._props.name, name)
        if inst._props.bitmap:
            set_prop(inst._props.bitmap, inst)
        set_prop(inst._props.mode, inst)
        set_prop(inst._props.offset_x, inst)
        set_prop(inst._props.offset_y, inst)
        set_prop(inst._props.pos_x, inst)
        set_prop(inst._props.pos_y, inst)
        set_prop(inst._props.point, inst)
        set_prop(inst._props.size_x, inst)
        set_prop(inst._props.size_y, inst)
        set_prop(inst._props.style, inst)

        # FillBitmapStretch and FillBitmapTile properties should not be used anymore
        # The FillBitmapMode property can be used instead to set all supported bitmap modes.
        # set_prop("FillBitmapStretch", inst)
        # set_prop("FillBitmapTile", inst)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()
    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.TXT_CONTENT | FormatKind.FILL
        return self._format_kind_prop

    @property
    def prop_bitmap(self) -> XBitmap | None:
        """Gets bitmap"""
        if not self._props.bitmap:
            return None
        pv = self._get(self._props.bitmap)
        return None if pv is None else pv

    @property
    def prop_mode(self) -> ImgStyleKind | None:
        """Gets/Sets if fill image is tiled"""
        pv = self._get(self._props.mode)
        return None if pv is None else ImgStyleKind.from_bitmap_mode(pv)

    @prop_mode.setter
    def prop_mode(self, value: ImgStyleKind | None) -> None:
        if value is None:
            self._remove(self._props.mode)
            return
        self._set(self._props.mode, value.value)

    @property
    def prop_is_size_percent(self) -> bool:
        """Gets if size is stored in percentage units."""
        x = cast(int, self._get(self._props.size_x))
        y = cast(int, self._get(self._props.size_y))
        return False if x is None or y is None else x < 0 or y < 0

    @property
    def prop_is_size_mm(self) -> bool:
        """Gets if size is stored in ``mm`` units."""
        x = cast(int, self._get(self._props.size_x))
        y = cast(int, self._get(self._props.size_y))
        return False if x is None or y is None else x >= 0 and y >= 0

    @property
    def prop_size(self) -> SizePercent | SizeMM | None:
        """Gets/Sets if fill image is stretched"""
        # size percent is stored as negative int
        # size mm is stored as 1/100th mm as positive int

        x = cast(int, self._get(self._props.size_x))
        y = cast(int, self._get(self._props.size_y))
        if x is None or y is None:
            return None
        if x == 0 and y == 0:
            return SizeMM(0, 0)
        if x < 0 or y < 0:
            # percent
            return SizePercent(round(abs(x)), round(abs(y)))
        x_val = UnitConvert.convert_mm100_mm(x)
        y_val = UnitConvert.convert_mm100_mm(y)
        return SizeMM(x_val, y_val)

    @prop_size.setter
    def prop_size(self, value: SizePercent | SizeMM | None) -> None:
        if value is None:
            self._remove(self._props.size_x)
            self._remove(self._props.size_y)
            return

        if isinstance(value, SizePercent):
            self._set(self._props.size_x, -value.width)
            self._set(self._props.size_y, -value.height)
            return
        if isinstance(value, SizeMM):
            x = UnitConvert.convert_mm_mm100(value.height)
            y = UnitConvert.convert_mm_mm100(value.width)
            self._set(self._props.size_x, x)
            self._set(self._props.size_y, y)
            return

    @property
    def prop_position(self) -> RectanglePoint | None:
        """Gets/Sets if fill image is tiled"""
        return self._get(self._props.point)

    @prop_position.setter
    def prop_position(self, value: RectanglePoint | None) -> None:
        if value is None:
            self._remove(self._props.point)
            return
        self._set(self._props.point, value)

    @property
    def prop_pos_offset(self) -> Offset | None:
        """Gets/Sets Position Offset"""
        x = cast(int, self._get(self._props.pos_x))
        y = cast(int, self._get(self._props.pos_y))
        return None if x is None or y is None else Offset(x, y)

    @prop_pos_offset.setter
    def prop_pos_offset(self, value: Offset | None) -> None:
        if value is None:
            self._remove(self._props.pos_x)
            self._remove(self._props.pos_y)
            return
        self._set(self._props.pos_x, value.x)
        self._set(self._props.pos_y, value.y)

    @property
    def prop_is_offset_row(self) -> bool:
        """Gets if the offset value is a row offset."""
        row = cast(int, self._get(self._props.offset_x))
        col = cast(int, self._get(self._props.offset_y))
        return False if row is None or col is None else col <= 0

    @property
    def prop_is_offset_column(self) -> bool:
        """Gets if the offset value is a column offset."""
        col = cast(int, self._get(self._props.offset_y))
        return False if col is None else col > 0

    @property
    def prop_tile_offset(self) -> OffsetColumn | OffsetRow | None:
        """Gets/Sets Tile Offset"""
        # row is X
        # col is Y
        row = cast(int, self._get(self._props.offset_x))
        col = cast(int, self._get(self._props.offset_y))
        if row is None or col is None:
            return None
        return OffsetColumn(col) if col > 0 else OffsetRow(row)

    @prop_tile_offset.setter
    def prop_tile_offset(self, value: OffsetColumn | OffsetRow | None) -> None:
        if value is None:
            self._remove(self._props.offset_x)
            self._remove(self._props.offset_y)
            return
        if isinstance(value, OffsetColumn):
            self._set(self._props.offset_x, 0)
            self._set(self._props.offset_y, value.value)
        else:
            self._set(self._props.offset_x, value.value)
            self._set(self._props.offset_y, 0)

    @property
    def _props(self) -> AreaImgProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaImgProps(
                name="FillBitmapName",
                style="FillStyle",
                mode="FillBitmapMode",
                point="FillBitmapRectanglePoint",
                bitmap="FillBitmap",
                offset_x="FillBitmapOffsetX",
                offset_y="FillBitmapOffsetY",
                pos_x="FillBitmapPositionOffsetX",
                pos_y="FillBitmapPositionOffsetY",
                size_x="FillBitmapSizeX",
                size_y="FillBitmapSizeY",
            )
        return self._props_internal_attributes

    # endregion Properties
