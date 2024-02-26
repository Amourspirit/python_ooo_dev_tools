from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooodev.format.inner.direct.chart2.chart.borders.line_properties import LineProperties as LineProps
from ooodev.format.inner.preset.preset_border_line import BorderLineKind
from ooodev.format.inner.preset.preset_border_line import get_preset_border_line_props
from ooodev.utils.color import Color
from ooodev.utils.data_type.intensity import Intensity

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class LineProperties(LineProps):
    """
    This class represents the line properties of a shape.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_line_properties`

    .. versionadded:: 0.17.4
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
            - :ref:`help_draw_format_direct_shape_line_properties`
        """
        super().__init__(style=style, color=color, width=width, transparency=transparency)

    def _set_line_style(self, style: BorderLineKind):
        # LineCaps and LineJoint are set Separately using CornerCap
        props = get_preset_border_line_props(kind=style)
        # self._set("LineCap", props.line_cap)
        self._set("LineDash", props.line_dash)
        self._set("LineDashName", props.line_dash_name)
        # self._set("LineJoint", props.line_joint)
        self._set("LineStyle", props.line_style)
