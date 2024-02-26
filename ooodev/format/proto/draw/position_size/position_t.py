from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.chart2.position_size.position_t import PositionT as ChartPositionT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind
else:
    Protocol = object
    UnitT = Any
    ShapeBasePointKind = Any


class PositionT(ChartPositionT, Protocol):
    """Fill Image Protocol"""

    def __init__(
        self,
        *,
        pos_x: float | UnitT,
        pos_y: float | UnitT,
        base_point: ShapeBasePointKind = ...,
    ) -> None:
        """
        Constructor

        Args:
            pos_x (float | UnitT): Specifies the x-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            pos_y (float | UnitT): Specifies the y-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            base_point (ShapeBasePointKind): Specifies the base point of the shape used to calculate the X and Y coordinates. Default is ``TOP_LEFT``.

        Returns:
            None:
        """

        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> PositionT:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            PositionT: New instance.
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> PositionT:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.
            **kwargs: Additional arguments.

        Returns:
            PositionT: New instance.
        """
        ...

    # region Properties

    @property
    def prop_base_point(self) -> ShapeBasePointKind:
        """
        Gets/Sets the base point of the shape used to calculate the X and Y coordinates.

        Returns:
            ShapeBasePointKind: Base point.
        """
        ...

    @prop_base_point.setter
    def prop_base_point(self, value: ShapeBasePointKind) -> None: ...

    # endregion Properties
