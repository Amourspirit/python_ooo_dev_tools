from __future__ import annotations
from typing import Any, Tuple
import uno
from com.sun.star.awt import XBitmap
from com.sun.star.chart2 import XChartDocument

from ooodev.loader import lo as mLo
from ooodev.format.inner.direct.chart2.chart.area.pattern import Pattern as ChartAreaPattern


class Pattern(ChartAreaPattern):
    """
    Class for Chart Wall Area Fill Pattern.

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
            - :ref:`help_chart2_format_direct_wall_floor_area`
        """
        super().__init__(
            chart_doc=chart_doc, bitmap=bitmap, name=name, tile=tile, stretch=stretch, auto_name=auto_name
        )

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet")
