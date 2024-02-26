from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.format.inner.preset.preset_border_line import BorderLineKind
from ooodev.utils.color import Color
from ooodev.utils.data_type.intensity import Intensity

from ooodev.format.inner.direct.chart2.chart.borders.line_properties import LineProperties as ChartLineProperties

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class LineProperties(ChartLineProperties):
    """
    This class represents the line properties of a chart axis line properties.

    .. seealso::

        - :ref:`help_chart2_format_direct_axis_line`
    """

    def __init__(
        self,
        style: BorderLineKind = BorderLineKind.CONTINUOUS,
        color: Color = Color(0),
        width: float | UnitT = 0,
        transparency: int | Intensity = 0,
    ) -> None:
        """
        Constructor.

        Args:
            style (BorderLineKind): Line style. Defaults to ``BorderLineKind.CONTINUOUS``.
            color (Color, optional): Line Color. Defaults to ``Color(0)``.
            width (float | UnitT, optional): Line Width (in ``mm`` units) or :ref:`proto_unit_obj`. Defaults to ``0``.
            transparency (int | Intensity, optional): Line transparency from ``0`` to ``100``. Defaults to ``0``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_axis_line`
        """
        super().__init__(style=style, color=color, width=width, transparency=transparency)
