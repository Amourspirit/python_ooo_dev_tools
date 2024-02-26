from __future__ import annotations
from typing import Any, Tuple
import uno
from com.sun.star.chart2 import XChartDocument
from ooo.dyn.drawing.hatch_style import HatchStyle

from ooodev.format.inner.direct.chart2.chart.area.hatch import Hatch as ChartAreaHatch
from ooodev.loader import lo as mLo
from ooodev.units.angle import Angle
from ooodev.units.unit_obj import UnitT
from ooodev.utils.color import Color


class Hatch(ChartAreaHatch):
    """
    Class for Chart Wall/Floor Area Fill Hatch.

    .. seealso::

        - :ref:`help_chart2_format_direct_wall_floor_area`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        style: HatchStyle = HatchStyle.SINGLE,
        color: Color = Color(0),
        space: float | UnitT = 0,
        angle: Angle | int = 0,
        bg_color: Color = Color(-1),
    ) -> None:
        """
        Constructor.

        Args:
            chart_doc (XChartDocument): Chart document.
            style (HatchStyle, optional): Specifies the kind of lines used to draw this hatch. Default ``HatchStyle.SINGLE``.
            color (:py:data:`~.utils.color.Color`, optional): Specifies the color of the hatch lines. Default ``0``.
            space (float, UnitT, optional): Specifies the space between the lines in the hatch (in ``mm`` units) or :ref:`proto_unit_obj`. Default ``0.0``
            angle (Angle, int, optional): Specifies angle of the hatch in degrees. Default to ``0``.
            bg_color(Color, optional): Specifies the background Color. Set this ``-1`` (default) for no background color.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_wall_floor_area`
        """
        super().__init__(chart_doc=chart_doc, style=style, color=color, space=space, angle=angle, bg_color=bg_color)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet")
