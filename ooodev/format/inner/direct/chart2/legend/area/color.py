from __future__ import annotations

from ooodev.utils import color as mColor
from ooodev.format.inner.direct.chart2.chart.area.color import Color as ChartColor


class Color(ChartColor):
    """
    Class for Chart Title Area Fill Color.

    .. seealso::

        - :ref:`help_chart2_format_direct_legend_area`

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
            - :ref:`help_chart2_format_direct_legend_area`
        """
        super().__init__(color=color)
