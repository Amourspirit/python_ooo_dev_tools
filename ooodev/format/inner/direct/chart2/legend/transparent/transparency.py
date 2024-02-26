from __future__ import annotations
import uno

from ooodev.utils.data_type.intensity import Intensity
from ooodev.format.inner.direct.chart2.chart.transparent.transparency import (
    Transparency as ChartTransparentTransparency,
)


class Transparency(ChartTransparentTransparency):
    """
    Chart Legend Transparent Transparency.

    ..seealso::

        - :ref:`help_chart2_format_direct_legend_transparency`

    .. versionadded:: 0.9.4
    """

    def __init__(self, value: Intensity | int = 0) -> None:
        """
        Constructor

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.

        Returns:
            None:

        See Also:

            - :ref:`help_chart2_format_direct_legend_transparency`
        """
        super().__init__(value=value)
