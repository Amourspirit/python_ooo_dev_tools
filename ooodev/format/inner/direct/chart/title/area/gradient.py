# region Import
from __future__ import annotations
from typing import Tuple
import uno

from ooodev.format.inner.direct.write.fill.area.gradient import Gradient as FillGradient

# endregion Import


class Gradient(FillGradient):
    """
    Style for Chart Title Area Fill Gradient.

    .. versionadded:: 0.9.4
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.Title",)
        return self._supported_services_values
