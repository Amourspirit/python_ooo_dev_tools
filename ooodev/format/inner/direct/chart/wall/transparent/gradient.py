"""
Module for Fill Gradient Color.

.. versionadded:: 0.9.0
"""
# region Import
from __future__ import annotations
import uno


from typing import Tuple

from ooodev.utils import lo as mLo
from ...chart.transparent.gradient import Gradient as ChartTransparentGradient

# endregion Import


class Gradient(ChartTransparentGradient):
    """
    Chart Wall Fill Gradient Color

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
