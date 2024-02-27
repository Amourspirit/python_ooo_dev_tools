from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.utils.kind.graphic_arrow_style_kind import GraphicArrowStyleKind
from ooodev.format.inner.direct.draw.shape.line.arrow_styles import ArrowStyles as DrawShapeLineArrowStyles

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class ArrowStyles(DrawShapeLineArrowStyles):
    """
    This class represents the line Arrow styles of a shape or line.

    .. seealso::

        - :ref:`help_draw_format_direct_shape_line_arrow_styles`

    .. versionadded:: 0.17.4
    """

    def __init__(
        self,
        start_line_name: GraphicArrowStyleKind | str | None = None,
        start_line_width: float | UnitT | None = None,
        start_line_center: bool | None = None,
        end_line_name: GraphicArrowStyleKind | str | None = None,
        end_line_width: float | UnitT | None = None,
        end_line_center: bool | None = None,
    ) -> None:
        """
        Constructor.

        Args:
            start_line_name (GraphicArrowStyleKind, str, optional): Start line name. Defaults to ``None``.
            start_line_width (float, UnitT, optional): Start line width in mm units. Defaults to ``None``.
            start_line_center (bool, optional): Start line center. Defaults to ``None``.
            end_line_name (GraphicArrowStyleKind, str, optional): End line name. Defaults to ``None``.
            end_line_width (float, UnitT, optional): End line width in mm units. Defaults to ``None``.
            end_line_center (bool, optional): End line center. Defaults to ``None``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_direct_shape_line_arrow_styles`
        """
        super().__init__(
            start_line_name=start_line_name,
            start_line_width=start_line_width,
            start_line_center=start_line_center,
            end_line_name=end_line_name,
            end_line_width=end_line_width,
            end_line_center=end_line_center,
        )
