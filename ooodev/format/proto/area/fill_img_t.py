from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno


from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol

    from com.sun.star.awt import XBitmap
    from ooo.dyn.drawing.rectangle_point import RectanglePoint
    from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
    from ooodev.format.inner.common.format_types.offset_row import OffsetRow
    from ooodev.format.inner.common.format_types.size_percent import SizePercent
    from ooodev.format.inner.preset.preset_image import PresetImageKind
    from ooodev.utils.data_type.offset import Offset
    from ooodev.utils.data_type.size_mm import SizeMM
    from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
else:
    Protocol = object
    XBitmap = Any
    RectanglePoint = Any
    OffsetColumn = Any
    OffsetRow = Any
    SizePercent = Any
    PresetImageKind = Any
    Offset = Any
    SizeMM = Any
    ImgStyleKind = Any
    


class FillImgT(StyleT, Protocol):
    """Fill Image Protocol"""

    def __init__(
        self,
        *,
        bitmap: XBitmap | None = ...,
        name: str = ...,
        mode: ImgStyleKind = ...,
        size: SizePercent | SizeMM | None = ...,
        position: RectanglePoint | None = ...,
        pos_offset: Offset | None = ...,
        tile_offset: OffsetColumn | OffsetRow | None = ...,
        auto_name: bool = ...,
    ) -> None:

        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> FillImgT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> FillImgT: ...

    @overload
    @classmethod
    def from_preset(cls, preset: PresetImageKind) -> FillImgT: ...

    @overload
    @classmethod
    def from_preset(cls, preset: PresetImageKind, **kwargs) -> FillImgT: ...

    # region Properties
    @property
    def prop_bitmap(self) -> XBitmap | None:
        """Gets bitmap"""
        ...

    @property
    def prop_mode(self) -> ImgStyleKind | None:
        """Gets/Sets if fill image is tiled"""
        ...

    @prop_mode.setter
    def prop_mode(self, value: ImgStyleKind | None) -> None:
        ...

    @property
    def prop_is_size_percent(self) -> bool:
        """Gets if size is stored in percentage units."""
        ...

    @property
    def prop_is_size_mm(self) -> bool:
        """Gets if size is stored in ``mm`` units."""
        ...

    @property
    def prop_size(self) -> SizePercent | SizeMM | None:
        """Gets/Sets if fill image is stretched"""
        ...

    @prop_size.setter
    def prop_size(self, value: SizePercent | SizeMM | None) -> None:
        ...

    @property
    def prop_position(self) -> RectanglePoint | None:
        """Gets/Sets if fill image is tiled"""
        ...

    @prop_position.setter
    def prop_position(self, value: RectanglePoint | None) -> None:
        ...

    @property
    def prop_pos_offset(self) -> Offset | None:
        """Gets/Sets Position Offset"""
        ...

    @prop_pos_offset.setter
    def prop_pos_offset(self, value: Offset | None) -> None:
        ...

    @property
    def prop_is_offset_row(self) -> bool:
        """Gets if the offset value is a row offset."""
        ...

    @property
    def prop_is_offset_column(self) -> bool:
        """Gets if the offset value is a column offset."""
        ...

    @property
    def prop_tile_offset(self) -> OffsetColumn | OffsetRow | None:
        """Gets/Sets Tile Offset"""
        ...

    @prop_tile_offset.setter
    def prop_tile_offset(self, value: OffsetColumn | OffsetRow | None) -> None:
        ...


    # endregion Properties
