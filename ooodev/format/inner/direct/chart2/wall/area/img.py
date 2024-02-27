# region Imports
from __future__ import annotations
from typing import Any, Tuple
import uno
from com.sun.star.awt import XBitmap
from com.sun.star.chart2 import XChartDocument
from ooo.dyn.drawing.rectangle_point import RectanglePoint

from ooodev.format.inner.common.format_types.offset_column import OffsetColumn
from ooodev.format.inner.common.format_types.offset_row import OffsetRow
from ooodev.format.inner.common.format_types.size_percent import SizePercent
from ooodev.format.inner.direct.write.fill.area.img import ImgStyleKind
from ooodev.loader import lo as mLo
from ooodev.utils.data_type.offset import Offset
from ooodev.utils.data_type.size_mm import SizeMM
from ooodev.format.inner.direct.chart2.chart.area.img import Img as ChartAreaImg

# endregion Imports


class Img(ChartAreaImg):
    """
    Class for Chart Wall Area Fill Image.

    .. seealso::

        - :ref:`help_chart2_format_direct_wall_floor_area`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        chart_doc: XChartDocument,
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
            chart_doc (XChartDocument): Chart document.
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

        Returns:
            None:

        Note:
            If ``auto_name`` is ``False`` then a bitmap for a given ``name`` is only required the first call.
            All subsequent call of the same ``name`` will retrieve the bitmap form the LibreOffice Bitmap Table.

        See Also:
            - :ref:`help_chart2_format_direct_wall_floor_area`
        """
        super().__init__(
            chart_doc=chart_doc,
            bitmap=bitmap,
            name=name,
            mode=mode,
            size=size,
            position=position,
            pos_offset=pos_offset,
            tile_offset=tile_offset,
            auto_name=auto_name,
        )

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet")
