from __future__ import annotations
from typing import Tuple

from ooodev.format.inner.direct.write.fill.area.fill_color import FillColor


class Color(FillColor):
    """
    Class for Chart Area Fill Color.

    .. versionadded:: 0.9.4
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.PageBackground",
                "com.sun.star.chart2.Title",
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.DataPoint",
            )
        return self._supported_services_values

    def _is_valid_obj(self, obj: object) -> bool:
        return self._is_obj_service(obj)
