from __future__ import annotations
import uno
from com.sun.star.awt import XBitmap
from com.sun.star.chart2 import XChartDocument

from ooodev.format.inner.direct.chart2.chart.area.pattern import Pattern as ChartPattern


class Pattern(ChartPattern):
    """
    Class for Chart Title Area Fill Pattern.

    .. seealso::

        - :ref:`help_chart2_format_direct_legend_area`

    .. versionadded:: 0.9.4
    """

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
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_legend_area`
        """
        super().__init__(
            chart_doc=chart_doc, bitmap=bitmap, name=name, tile=tile, stretch=stretch, auto_name=auto_name
        )
