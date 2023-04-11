from __future__ import annotations
import uno
from typing import Tuple

from ooodev.utils import lo as mLo
from ...chart.borders.line_properties import LineProperties as ChartBordersLineProperties


class LineProperties(ChartBordersLineProperties):
    """This class represents the line properties of a chart wall borders line properties."""

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _is_valid_obj(self, obj: object) -> bool:
        return mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet")
