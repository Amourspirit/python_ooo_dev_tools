"""
Module for Fill Properties Fill Image.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from dataclasses import dataclass
from typing import Any, Tuple, cast, overload, Type, TypeVar, TYPE_CHECKING
from enum import Enum

import uno

from .....events.args.key_val_cancel_args import KeyValCancelArgs
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.data_type.intensity import Intensity
from .....utils.data_type.offset import Offset as Offset
from .....utils.data_type.width_height_fraction import WidthHeightFraction
from .....utils.data_type.width_height_percent import WidthHeightPercent
from .....utils.unit_convert import UnitConvert
from ....kind.format_kind import FormatKind
from ....preset import preset_image as mImage
from ....preset.preset_image import PresetImageKind as PresetImageKind
from ....style_base import StyleBase


from com.sun.star.awt import XBitmap

from ooo.dyn.drawing.fill_style import FillStyle as FillStyle
from ooo.dyn.drawing.bitmap_mode import BitmapMode
from ooo.dyn.drawing.rectangle_point import RectanglePoint as RectanglePoint

if TYPE_CHECKING:
    from com.sun.star.graphic import Graphic

# https://github.com/LibreOffice/core/blob/6379414ca34527fbe69df2035d49d651655317cd/vcl/source/filter/ipict/ipict.cxx#L92

_TImg = TypeVar(name="_TImg", bound="Img")


@dataclass(frozen=True)
class SizeMM(WidthHeightFraction):
    """Size in ``mm`` units"""

    pass


@dataclass(frozen=True)
class SizePercent(WidthHeightPercent):
    """Size in percent values"""

    pass


@dataclass(unsafe_hash=True)
class OffsetColumn(Intensity):
    """Represents a Column Offset value from ``0`` to ``100``."""

    pass


@dataclass(unsafe_hash=True)
class OffsetRow(Intensity):
    """Represents a Row Offset value from ``0`` to ``100``."""

    pass


class ImgStyleKind(Enum):
    """Image Style Kind for Image Fill"""

    CUSTOM = BitmapMode.NO_REPEAT
    """Image does not repeate"""
    TILED = BitmapMode.REPEAT
    """Image Repeates"""
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

        # when mode is ImgStyleKind.STRETCHED size, position, pos_offset, and tile_offset are not required

        init_vals = {}
        bmap = None
        try:
            # if bitmap or name is passed in then get the bitmap
            bmap = self._get_bitmap(bitmap, name, auto_name)
        except Exception:
            pass
        if not bmap is None:

            init_vals["FillBitmap"] = bmap
            init_vals["FillBitmapName"] = self._name
            init_vals["FillStyle"] = FillStyle.BITMAP

        super().__init__(**init_vals)
        self.prop_mode = mode
        self.prop_size = size
        self.prop_posiion = position
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
        bmap = self._container_get_value(self._name, nc)  # raises value error if name is empty
        if not bmap is None:
            return bmap
        if bitmap is None:
            raise ValueError(f'No bitmap could be found in container for "{name}". In this case a bitmap is required.')
        self._container_add_value(name=self._name, obj=bitmap, allow_update=False, nc=nc)
        return self._container_get_value(self._name, nc)

    # endregion Internal Methods

    # region Overrides

    def copy(self: _TImg) -> _TImg:
        cp = super().copy()
        cp._name = self._name
        return cp

    def _container_get_service_name(self) -> str:
        # https://github.com/LibreOffice/core/blob/d9e044f04ac11b76b9a3dac575f4e9155b67490e/chart2/source/tools/PropertyHelper.cxx#L246
        return "com.sun.star.drawing.BitmapTable"

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.text.TextContent",
            "com.sun.star.beans.PropertySet",
            "com.sun.star.style.ParagraphStyle",
        )

    # region apply()
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
        if not self._has("FillBitmap"):
            mLo.Lo.print("Img.apply(): There is nothing to apply.")
            return
        super().apply(obj, **kwargs)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"Img.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    def on_property_restore_setting(self, event_args: KeyValCancelArgs) -> None:
        if event_args.key == "FillBitmap":
            if event_args.value is None:
                event_args.default = True
        elif event_args.key == "FillBitmapName":
            if not event_args.value:
                event_args.default = True
        return super().on_property_restore_setting(event_args)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_preset(cls: Type[_TImg], preset: PresetImageKind) -> _TImg:
        """
        Gets an instance from a preset

        Args:
            preset (ImageKind): Preset

        Returns:
            Img: Instance from preset.
        """
        name = str(preset)
        nu = super(Img, cls).__new__(cls)
        nu.__init__()

        nc = nu._container_get_inst()
        bmap = cast("Graphic", nu._container_get_value(name, nc))
        if bmap is None:
            bmap = mImage.get_prest_bitmap(preset)
        inst = super(Img, cls).__new__(cls)
        inst.__init__(
            bitmap=bmap,
            name=name,
            mode=ImgStyleKind.TILED,
            position=RectanglePoint.MIDDLE_MIDDLE,
            pos_offset=Offset(0, 0),
            tile_offset=OffsetRow(0),
            auto_name=False,
        )
        # set size
        point = preset._get_point()
        inst._set("FillBitmapSizeX", point.x)
        inst._set("FillBitmapSizeY", point.y)
        return inst

    @classmethod
    def from_obj(cls: Type[_TImg], obj: object) -> _TImg:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Img: ``Img`` instance that represents ``obj`` fill image.
        """
        nu = super(Img, cls).__new__(cls)
        nu.__init__()
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not support to convert to Img")

        def set_prop(key: str, fp: Img):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                fp._set(key, val)

        inst = super(Img, cls).__new__(cls)
        inst.__init__()
        name = mProps.Props.get(obj, "FillBitmapName")
        inst._name = name
        inst._set("FillBitmapName", name)
        set_prop("FillBitmap", inst)
        set_prop("FillBitmapMode", inst)
        set_prop("FillBitmapOffsetX", inst)
        set_prop("FillBitmapOffsetY", inst)
        set_prop("FillBitmapPositionOffsetX", inst)
        set_prop("FillBitmapPositionOffsetY", inst)
        set_prop("FillBitmapRectanglePoint", inst)
        set_prop("FillBitmapSizeX", inst)
        set_prop("FillBitmapSizeY", inst)
        set_prop("FillStyle", inst)

        # FillBitmapStretch and FillBitmapTile properties should not be used anymore
        # The FillBitmapMode property can be used instead to set all supported bitmap modes.
        # set_prop("FillBitmapStretch", inst)
        # set_prop("FillBitmapTile", inst)
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.TXT_CONTENT | FormatKind.FILL

    @property
    def prop_mode(self) -> ImgStyleKind | None:
        """Gets/Sets if fill image is tiled"""
        pv = self._get("FillBitmapMode")
        if pv is None:
            return None
        return ImgStyleKind.from_bitmap_mode(pv)

    @prop_mode.setter
    def prop_mode(self, value: ImgStyleKind | None) -> None:
        if value is None:
            self._remove("FillBitmapMode")
            return
        self._set("FillBitmapMode", value.value)

    @property
    def prop_is_size_percent(self) -> bool:
        """Gets if size is stored in percentage units."""
        x = cast(int, self._get("FillBitmapSizeX"))
        y = cast(int, self._get("FillBitmapSizeY"))
        if x is None or y is None:
            return False
        if x < 0 or y < 0:
            return True
        return False

    @property
    def prop_is_size_mm(self) -> bool:
        """Gets if size is stored in ``mm`` units."""
        x = cast(int, self._get("FillBitmapSizeX"))
        y = cast(int, self._get("FillBitmapSizeY"))
        if x is None or y is None:
            return False
        if x < 0 or y < 0:
            return False
        return True

    @property
    def prop_size(self) -> SizePercent | SizeMM | None:
        """Gets/Sets if fill image is stretched"""
        # size percent is stored as negative int
        # size mm is stored as 1/100th mm as postitive int

        x = cast(int, self._get("FillBitmapSizeX"))
        y = cast(int, self._get("FillBitmapSizeY"))
        if x is None or y is None:
            return None
        if x == 0 and y == 0:
            return SizeMM(0, 0)
        if x < 0 or y < 0:
            # percent
            return SizePercent(round(abs(x)), round(abs(y)))
        xval = UnitConvert.convert_mm100_mm(x)
        yval = UnitConvert.convert_mm100_mm(y)
        return SizeMM(xval, yval)

    @prop_size.setter
    def prop_size(self, value: SizePercent | SizeMM | None) -> None:
        if value is None:
            self._remove("FillBitmapSizeX")
            self._remove("FillBitmapSizeY")
            return

        if isinstance(value, SizePercent):
            self._set("FillBitmapSizeX", -value.width)
            self._set("FillBitmapSizeY", -value.height)
            return
        if isinstance(value, SizeMM):
            x = UnitConvert.convert_mm_mm100(value.height)
            y = UnitConvert.convert_mm_mm100(value.width)
            self._set("FillBitmapSizeX", x)
            self._set("FillBitmapSizeY", y)
            return

    @property
    def prop_posiion(self) -> RectanglePoint | None:
        """Gets/Sets if fill image is tiled"""
        return self._get("FillBitmapRectanglePoint")

    @prop_posiion.setter
    def prop_posiion(self, value: RectanglePoint | None) -> None:
        if value is None:
            self._remove("FillBitmapRectanglePoint")
            return
        self._set("FillBitmapRectanglePoint", value)

    @property
    def prop_pos_offset(self) -> Offset | None:
        """Gets/Sets Position Offset"""
        x = cast(int, self._get("FillBitmapPositionOffsetX"))
        y = cast(int, self._get("FillBitmapPositionOffsetY"))
        if x is None or y is None:
            return None
        return Offset(x, y)

    @prop_pos_offset.setter
    def prop_pos_offset(self, value: Offset | None) -> None:
        if value is None:
            self._remove("FillBitmapPositionOffsetX")
            self._remove("FillBitmapPositionOffsetY")
            return
        self._set("FillBitmapPositionOffsetX", value.x)
        self._set("FillBitmapPositionOffsetY", value.y)

    @property
    def prop_is_offset_row(self) -> bool:
        """Gets if the offset value is a row offset."""
        row = cast(int, self._get("FillBitmapOffsetX"))
        col = cast(int, self._get("FillBitmapOffsetY"))
        if row is None or col is None:
            return False
        if col > 0:
            return False
        return True

    @property
    def prop_is_offset_column(self) -> bool:
        """Gets if the offset value is a column offset."""
        col = cast(int, self._get("FillBitmapOffsetY"))
        if col is None:
            return False
        return col > 0

    @property
    def prop_tile_offset(self) -> OffsetColumn | OffsetRow | None:
        """Gets/Sets Tile Offset"""
        # row is X
        # col is Y
        row = cast(int, self._get("FillBitmapOffsetX"))
        col = cast(int, self._get("FillBitmapOffsetY"))
        if row is None or col is None:
            return None
        if col > 0:
            return OffsetColumn(col)
        # if x and y are zero the default is to return row
        return OffsetRow(row)

    @prop_tile_offset.setter
    def prop_tile_offset(self, value: OffsetColumn | OffsetRow | None) -> None:
        if value is None:
            self._remove("FillBitmapOffsetX")
            self._remove("FillBitmapOffsetY")
            return
        if isinstance(value, OffsetColumn):
            self._set("FillBitmapOffsetX", 0)
            self._set("FillBitmapOffsetY", value.value)
        else:
            self._set("FillBitmapOffsetX", value.value)
            self._set("FillBitmapOffsetY", 0)

    # endregion Properties
