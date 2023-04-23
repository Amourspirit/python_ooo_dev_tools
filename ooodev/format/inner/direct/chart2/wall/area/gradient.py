# region Import
from __future__ import annotations
from typing import Tuple
import uno

from ooodev.utils import lo as mLo
from ...chart.area.gradient import Gradient as ChartAreaGradient

# endregion Import


class Gradient(ChartAreaGradient):
    """
    Style for Chart Wall Area Fill Gradient.

    .. versionadded:: 0.9.4
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _is_valid_obj(self, obj: object) -> bool:
        return mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet")
