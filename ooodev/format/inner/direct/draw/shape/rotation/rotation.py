from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.drawing import XShape
from ooo.dyn.awt.point import Point as UnoPoint

from ooodev.utils import lo as mLo
from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind
from ooodev.utils.data_type.size import Size
from ooodev.format.inner.direct.chart2.position_size.position import Position as ShapePosition
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.utils import info as mInfo
from ooodev.format.inner.style_base import StyleBase
from ooodev.utils.data_type.angle import Angle
from ..position_size import position as mPosition
from ooodev.utils.data_type.point import Point
from ooodev.utils.data_type.point_unit import PointUnit
from ooodev.units import UnitMM100

if TYPE_CHECKING:
    from ooodev.units import UnitT
    from com.sun.star.drawing import DrawPage
else:
    DrawPage = Any


class Rotation(StyleBase):
    """
    Rotation of a shape.

    .. versionadded:: 0.17.3
    """

    def __init__(
        self,
        rotation: int | Angle = 0,
        pivot_point: Point | PointUnit | None = None,
        base_point: ShapeBasePointKind = ShapeBasePointKind.CENTER,
    ) -> None:
        """
        Constructor

        Args:
            rotation (int, Angle, optional): Specifies the rotation angle of the shape in degrees.
                Default is ``0``.
            pivot_point (Point, PointUnit, optional): Specifies the x and y coordinates of the position of the shape.
                ``Point`` is in ``mm`` units and ``PointUnit`` is in :ref:`proto_unit_obj` units.
            base_point (ShapeBasePointKind): Specifies the base point of the shape used to calculate the X and Y coordinates. Default is ``CENTER``.

        Returns:
            None:
        """
        super().__init__()
        rotation = Angle(int(rotation))
        self._set("RotateAngle", rotation.value)
        self._pivot_point = None  # none or UnitMM100 points
        if pivot_point is not None:
            if isinstance(pivot_point, Point):  # mm
                self._pivot_point = PointUnit(UnitMM100.from_mm(pivot_point.x), UnitMM100.from_mm(pivot_point.y))
            else:
                self._pivot_point = PointUnit(
                    UnitMM100(pivot_point.x.get_value_mm100()), UnitMM100(pivot_point.y.get_value_mm100())
                )

        self._base_point = base_point

    # region override
    def apply(self, obj: Any, **kwargs) -> None:
        """
        Apply the rotation to the shape.

        Args:
            shape (XShape): Shape to apply the rotation to.

        Returns:
            None:
        """
        shape = mLo.Lo.qi(XShape, obj)
        if shape is None:
            return
        if self._pivot_point is not None:
            # must check the position
            # get the shapes current position
            pos = mPosition.Position.from_obj(shape)

            if (
                pos.prop_pos_x != self._pivot_point.x.get_value_mm100()
                or pos.prop_pos_y != self._pivot_point.y.get_value_mm100()
            ):
                raise mEx.Exc("Pivot point must be the same as the position of the shape")

        if self._base_point != ShapeBasePointKind.CENTER:
            self._set("RotateAngleReference", self._base_point.value)

        super().apply(shape)

    # endregion override

    # region Properties
    @property
    def prop_rotation(self) -> Angle:
        """
        Gets/Sets Rotation angle of the shape in degrees.

        Property can be set by passing int or Angle.

        Returns:
            Angle:
        """
        return Angle(self._get("RotateAngle"))

    @prop_rotation.setter
    def prop_rotation(self, value: int | Angle) -> None:
        self._set("RotateAngle", Angle(int(value)).value)

    # endregion Properties
