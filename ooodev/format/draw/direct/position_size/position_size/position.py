from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.format.inner.direct.draw.shape.position_size.position import Position as ShapePosition
from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class Position(ShapePosition):
    """
    Shape Position

    .. seealso::

        - :ref:`help_draw_format_direct_shape_position_size_position_size_position`

    .. versionadded:: 0.17.3
    """

    def __init__(
        self,
        pos_x: float | UnitT,
        pos_y: float | UnitT,
        base_point: ShapeBasePointKind = ShapeBasePointKind.TOP_LEFT,
    ) -> None:
        """
        Constructor

        Args:
            pos_x (float | UnitT): Specifies the x-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            pos_y (float | UnitT): Specifies the y-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            base_point (ShapeBasePointKind): Specifies the base point of the shape used to calculate the X and Y coordinates. Default is ``TOP_LEFT``.

        Returns:
            None:

        Note:
            ``pos_x`` and ``pos_y`` are the coordinates of the shape inside the draw page borders.
            This is the same behavior as the dialog box.
            If the draw page has a border of 10mm and the shape is positioned at 0mm,0mm in the dialog box then the shape
            is actually at 10mm,10mm relative to the draw page document.

        See Also:
            - :ref:`help_draw_format_direct_shape_position_size_position_size_position`
        """
        super().__init__(pos_x=pos_x, pos_y=pos_y, base_point=base_point)
