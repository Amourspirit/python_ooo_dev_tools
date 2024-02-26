from __future__ import annotations
import uno
from typing import Any, Tuple

from ooodev.format.inner.direct.chart2.chart.borders.line_properties import (
    LineProperties as ChartBordersLineProperties,
)
from ooodev.format.inner.preset.preset_border_line import BorderLineKind
from ooodev.loader import lo as mLo
from ooodev.units.unit_obj import UnitT
from ooodev.utils.color import Color
from ooodev.utils.data_type.intensity import Intensity


class LineProperties(ChartBordersLineProperties):
    """
    This class represents the line properties of a chart wall borders line properties.

    .. seealso::

        - :ref:`help_chart2_format_direct_wall_floor_borders`
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
            - :ref:`help_chart2_format_direct_wall_floor_borders`
        """
        super().__init__(style=style, color=color, width=width, transparency=transparency)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet")
