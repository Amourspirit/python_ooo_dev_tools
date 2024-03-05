from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno


from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_multi_t import StyleMultiT


if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol

    from com.sun.star.awt import XBitmap
    from ooo.dyn.drawing.rectangle_point import RectanglePoint
    from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
    from ooodev.format.inner.common.format_types.offset_row import OffsetRow
    from ooodev.format.inner.common.format_types.size_percent import SizePercent
    from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
    from ooodev.format.inner.preset.preset_image import PresetImageKind
    from ooodev.format.proto.area.fill_img_t import FillImgT as InnerFillImgT
    from ooodev.utils.data_type.offset import Offset
    from ooodev.utils.data_type.size_mm import SizeMM
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
    InnerFillImgT = Any
    


class FillImgT(StyleMultiT, Protocol):
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
    def prop_inner(self) -> InnerFillImgT:
        """Gets Fill image instance"""
        ...

    # endregion Properties
