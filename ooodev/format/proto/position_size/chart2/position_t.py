from __future__ import annotations
from typing import Any, TYPE_CHECKING


if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
    from ...style_t import StyleT

    from ooodev.units import UnitT, UnitMM
else:
    Protocol = object
    UnitT = Any
    UnitMM = Any


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
