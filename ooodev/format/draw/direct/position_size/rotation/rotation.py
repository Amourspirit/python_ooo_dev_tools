from __future__ import annotations
from typing import TYPE_CHECKING
import uno

if TYPE_CHECKING:
    from ooodev.units import AngleUnitT

from ooodev.format.inner.direct.draw.shape.rotation.rotation import Rotation as ShapeRotation


class Rotation(ShapeRotation):
    """
    Shape Rotation

    .. seealso::

        - :ref:`help_draw_format_direct_shape_position_size_position_rotation`

    .. versionadded:: 0.17.4
    """

    def __init__(
        self,
        rotation: int | AngleUnitT = 0,
    ) -> None:
        """
        Constructor

        Args:
            rotation (int, AngleUnitT, optional): Specifies the rotation angle of the shape in degrees.
                Default is ``0``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_direct_shape_position_size_position_rotation`
        """
        super().__init__(rotation=rotation)
