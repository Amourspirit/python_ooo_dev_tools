from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

# pylint: disable=ungrouped-imports
from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol

    from ooodev.units.unit_obj import UnitT
    from ooodev.units.unit_mm import UnitMM
else:
    Protocol = object
    UnitT = Any
    UnitMM = Any


# see ooodev.format.inner.direct.chart2.position_size.position.Position
class PositionT(StyleT, Protocol):
    """Position Protocol"""

    def __init__(
        self,
        *,
        pos_x: float | UnitT,
        pos_y: float | UnitT,
    ) -> None:
        """
        Constructor

        Args:
            pos_x (float, UnitT): Specifies the x-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            pos_y (float, UnitT): Specifies the y-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
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
    def prop_pos_x(self) -> UnitMM:
        """Gets or sets the x-coordinate of the position of the shape (in ``mm`` units)."""
        ...

    @prop_pos_x.setter
    def prop_pos_x(self, value: float | UnitT) -> None: ...

    @property
    def prop_pos_y(self) -> UnitMM:
        """Gets or sets the y-coordinate of the position of the shape (in ``mm`` units)."""
        ...

    @prop_pos_y.setter
    def prop_pos_y(self, value: float | UnitT) -> None: ...

    # endregion Properties
