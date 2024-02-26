from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.drawing.hatch_style import HatchStyle
from ooodev.utils.color import StandardColor

from ooodev.format.inner.direct.write.fill.area.hatch import Hatch as FillHatch

if TYPE_CHECKING:
    from ooodev.utils.color import Color
    from ooodev.units.angle import Angle
    from ooodev.units.unit_obj import UnitT


class Hatch(FillHatch):
    """
    Class for Fill Properties Fill Hatch.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_area_hatch`

    .. versionadded:: 0.9.3
    """

    def __init__(
        self,
        *,
        style: HatchStyle = HatchStyle.SINGLE,
        color: Color = StandardColor.BLACK,
        space: float | UnitT = 0.0,
        angle: Angle | int = 0,
        bg_color: Color = StandardColor.AUTO_COLOR,
    ) -> None:
        """
        Constructor

        Args:
            style (HatchStyle, optional): Specifies the kind of lines used to draw this hatch. Default ``HatchStyle.SINGLE``.
            color (:py:data:`~.utils.color.Color`, optional): Specifies the color of the hatch lines. Default ``0``.
            space (float, UnitT, optional): Specifies the space between the lines in the hatch (in ``mm`` units) or :ref:`proto_unit_obj`. Default ``0.0``
            angle (Angle, int, optional): Specifies angle of the hatch in degrees. Default to ``0``.
            bg_color(Color, optional): Specifies the background Color. Set this ``-1`` (default) for no background color.

        Returns:
            None:

        See Also:

            - :ref:`help_draw_format_direct_shape_area_hatch`
        """
        super().__init__(style=style, color=color, space=space, angle=angle, bg_color=bg_color)
