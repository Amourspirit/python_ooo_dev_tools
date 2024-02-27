from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.format.inner.direct.draw.shape.position_size.size import Size as ShapeSize
from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class Size(ShapeSize):
    """
    Shape Size

    .. seealso::

        - :ref:`help_draw_format_direct_shape_position_size_position_size_size`

    .. versionadded:: 0.17.3
    """

    def __init__(
        self,
        width: float | UnitT,
        height: float | UnitT,
        base_point: ShapeBasePointKind = ShapeBasePointKind.TOP_LEFT,
    ) -> None:
        """
        Constructor

        Args:
            width (float | UnitT): Specifies the width of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            height (float | UnitT): Specifies the height of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            base_point (ShapeBasePointKind): Specifies the base point of the shape used to calculate the X and Y coordinates. Default is ``TOP_LEFT``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_direct_shape_position_size_position_size_size`
        """
        super().__init__(width=width, height=height, base_point=base_point)
