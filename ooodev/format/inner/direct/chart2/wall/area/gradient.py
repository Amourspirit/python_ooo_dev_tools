# region Import
from __future__ import annotations
from typing import Any, Tuple
import uno
from com.sun.star.chart2 import XChartDocument

from ooo.dyn.awt.gradient_style import GradientStyle
from ooodev.loader import lo as mLo
from ooodev.utils.color import Color
from ooodev.units.angle import Angle
from ooodev.utils.data_type.color_range import ColorRange
from ooodev.utils.data_type.intensity import Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset
from ooodev.format.inner.direct.chart2.chart.area.gradient import Gradient as ChartAreaGradient

# endregion Import


class Gradient(ChartAreaGradient):
    """
    Style for Chart Wall/Floor Area Fill Gradient.

    .. seealso::

        - :ref:`help_chart2_format_direct_wall_floor_area`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        step_count: int = 0,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_color: ColorRange = ColorRange(Color(0), Color(16777215)),
        grad_intensity: IntensityRange = IntensityRange(100, 100),
        name: str = "",
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (Offset, int, optional): Specifies the X and Y coordinate, where the gradient begins.
                X is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT``
                style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to ``0``.
            border (int, optional): Specifies percent of the total width where just the start color is used.
                Defaults to ``0``.
            grad_color (ColorRange, optional): Specifies the color at the start point and stop point of the gradient.
                Defaults to ``ColorRange(Color(0), Color(16777215))``.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of the
                gradient. Defaults to ``IntensityRange(100, 100)``.
            name (str, optional): Specifies the Fill Gradient Name.

        Returns:
            None:

        See Also:

            - :ref:`help_chart2_format_direct_wall_floor_area`
        """
        super().__init__(
            chart_doc=chart_doc,
            style=style,
            step_count=step_count,
            offset=offset,
            angle=angle,
            border=border,
            grad_color=grad_color,
            grad_intensity=grad_intensity,
            name=name,
        )

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet")
