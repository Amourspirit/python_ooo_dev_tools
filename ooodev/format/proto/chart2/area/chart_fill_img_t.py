from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.area.fill_img_t import FillImgT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from com.sun.star.chart2 import XChartDocument
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
    XChartDocument = Any
    XBitmap = Any
    RectanglePoint = Any
    OffsetColumn = Any
    OffsetRow = Any
    SizePercent = Any
    PresetImageKind = Any
    Offset = Any
    SizeMM = Any
    ImgStyleKind = Any
    


class ChartFillImgT(FillImgT, Protocol):
    """Chart Fill Image Protocol for charts"""

    def __init__(
        self,
        chart_doc: XChartDocument,
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
    def from_obj(cls,  chart_doc: XChartDocument, obj: Any) -> ChartFillImgT: ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any, **kwargs) -> ChartFillImgT: ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetImageKind) -> ChartFillImgT: ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetImageKind, **kwargs) -> ChartFillImgT: ...
