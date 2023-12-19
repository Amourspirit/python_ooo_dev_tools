from ooodev.format.inner.direct.draw.shape.line.corner_caps import CornerCaps as ShapesCornerCaps
from ooo.dyn.drawing.line_joint import LineJoint
from ooo.dyn.drawing.line_cap import LineCap


class CornerCaps(ShapesCornerCaps):
    """
    This class represents the line properties of a chart borders line properties.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_line_corner_caps`

    .. versionadded:: 0.17.4
    """

    def __init__(self, corner_style: LineJoint = LineJoint.ROUND, cap_style: LineCap = LineCap.BUTT) -> None:
        """
        Constructor.

        Args:
            corner_style (LineJoint, optional): Corner style. Defaults to ``LineJoint.ROUND``.
            cap_style (LineCap, optional): Cap style. Defaults to ``LineCap.BUTT``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_direct_shape_line_corner_caps`
        """
        super().__init__(corner_style=corner_style, cap_style=cap_style)
