# region Import
from __future__ import annotations
from typing import Tuple, Any
import uno
from com.sun.star.chart2 import XChartDocument

from ooo.dyn.awt.gradient_style import GradientStyle
from ooodev.utils import lo as mLo
from ooodev.utils.data_type.angle import Angle
from ooodev.utils.data_type.intensity import Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset
from ...chart.transparent.gradient import Gradient as ChartTransparentGradient

# endregion Import


class Gradient(ChartTransparentGradient):
    """
    Chart Wall Fill Gradient Color

    .. seealso::

        - :ref:`help_chart2_format_direct_wall_floor_transparency`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_intensity: IntensityRange = IntensityRange(0, 0),
        **kwargs: Any,
    ) -> None:
        """
        Constructor

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (offset, optional): Specifies the X-coordinate (start) and Y-coordinate (end),
                where the gradient begins. X is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and
                ``RECT`` style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to ``0``.
            border (int, optional): Specifies percent of the total width where just the start color is used.
                Defaults to ``0``.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of
                the gradient. Defaults to ``IntensityRange(0, 0)``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_wall_floor_transparency`
        """
        super().__init__(
            chart_doc=chart_doc,
            style=style,
            offset=offset,
            angle=angle,
            border=border,
            grad_intensity=grad_intensity,
            **kwargs,
        )

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return mLo.Lo.is_uno_interfaces(obj, "com.sun.star.beans.XPropertySet")
