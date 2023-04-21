from __future__ import annotations
from typing import Tuple

from ooodev.format.inner.direct.chart2.series.data_labels.data_labels.number_format import (
    NumberFormat as DataLabelsNumberFormat,
)


class Numbers(DataLabelsNumberFormat):
    """
    Chart Axis Numbers format.

    .. versionadded:: 0.9.4
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.Axis",)
        return self._supported_services_values
