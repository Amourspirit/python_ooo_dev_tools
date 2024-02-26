from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooodev.format.inner.direct.write.fill.transparent.gradient import Gradient as FillGradient

from ooo.dyn.awt.gradient_style import GradientStyle
from ooodev.utils.data_type.offset import Offset
from ooodev.utils.data_type.intensity_range import IntensityRange

if TYPE_CHECKING:
    from ooodev.units.angle import Angle
    from ooodev.utils.data_type.intensity import Intensity


class Gradient(FillGradient):
    """
    Fill Gradient Color

    .. seealso::

        - :ref:`help_writer_format_direct_shape_transparency_gradient`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
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

            - :ref:`help_writer_format_direct_shape_transparency_gradient`
        """
        super().__init__(
            style=style, offset=offset, angle=angle, border=border, grad_intensity=grad_intensity, **kwargs
        )
