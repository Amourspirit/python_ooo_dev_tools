# region Import
from __future__ import annotations
from typing import Any
import uno
from com.sun.star.chart2 import XChartDocument

from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle

from ooodev.utils.data_type.angle import Angle as Angle
from ooodev.utils.data_type.intensity import Intensity as Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange as IntensityRange
from ooodev.utils.data_type.offset import Offset as Offset


from ...chart.transparent.gradient import Gradient as ChartTransparentGradient

# endregion Import


class Gradient(ChartTransparentGradient):
    """
    Chart Legend Transparent Fill Gradient Color

    .. seealso::

        - :ref:`help_chart2_format_direct_legend_transparency`

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
            - :ref:`help_chart2_format_direct_legend_transparency`
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
