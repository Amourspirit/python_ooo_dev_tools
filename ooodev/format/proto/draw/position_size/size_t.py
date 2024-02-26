from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.chart2.position_size.size_t import SizeT as Chart2SizeT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind
else:
    Protocol = object
    UnitT = Any
    ShapeBasePointKind = Any


class SizeT(Chart2SizeT, Protocol):
    """Size Protocol"""

    def __init__(
        self,
        *,
        width: float | UnitT,
        height: float | UnitT,
        base_point: ShapeBasePointKind = ...,
    ) -> None:
        """
        Constructor

        Args:
            width (float | UnitT): Specifies the width of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            height (float | UnitT): Specifies the height of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            base_point (ShapeBasePointKind): Specifies the base point of the shape used to calculate the X and Y coordinates. Default is ``TOP_LEFT``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> SizeT:
        """
        Gets size from ``obj``

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Size: New instance.
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> SizeT:
        """
        Gets size from ``obj``

        Args:
            obj (Any): UNO Shape object.
            **kwargs: Additional arguments.

        Returns:
            Size: New instance.
        """
        ...

    # region properties
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

    # endregion properties
