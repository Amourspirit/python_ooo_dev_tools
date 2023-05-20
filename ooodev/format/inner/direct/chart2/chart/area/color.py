from __future__ import annotations
from typing import Any, Tuple

from ooodev.format.inner.direct.write.fill.area.fill_color import FillColor
from ooodev.utils import color as mColor


class Color(FillColor):
    """
    Class for Chart Area Fill Color.

    .. seealso::

        - :ref:`help_chart2_format_direct_general_area`

    .. versionadded:: 0.9.4
    """

    def __init__(self, color: mColor.Color = mColor.Color(-1)) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): Fill Color Color.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_general_area`
        """
        super().__init__(color=color)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.DataPoint",
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.Legend",
                "com.sun.star.chart2.PageBackground",
                "com.sun.star.chart2.Title",
                "com.sun.star.drawing.FillProperties",
            )
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return self._is_obj_service(obj)
