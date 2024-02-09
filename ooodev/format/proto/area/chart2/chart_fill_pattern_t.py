from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
    from com.sun.star.chart2 import XChartDocument
    from com.sun.star.awt import XBitmap
    from ooodev.format.inner.preset.preset_pattern import PresetPatternKind
    from ..fill_pattern_t import FillPatternT

else:
    Protocol = object
    PresetPatternKind = Any
    XBitmap = Any
    XChartDocument = Any



class ChartFillPatternT(FillPatternT, Protocol):
    """Fill Pattern Protocol"""

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        tile: bool = True,
        stretch: bool = False,
        auto_name: bool = False,
    ) -> None:

        ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any) -> ChartFillPatternT: ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument,obj: Any, **kwargs) -> ChartFillPatternT: ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetPatternKind) -> ChartFillPatternT: ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetPatternKind, **kwargs) -> ChartFillPatternT: ...
