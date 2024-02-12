from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol

    from ooodev.units import UnitT, UnitMM
else:
    Protocol = object
    UnitT = Any
    UnitMM = Any


class SizeT(StyleT, Protocol):
    """Size Protocol"""

    def __init__(
        self,
        *,
        width: float | UnitT,
        height: float | UnitT,
    ) -> None:
        """
        Constructor

        Args:
            width (float | UnitT): Specifies the width of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            height (float | UnitT): Specifies the height of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            None:
        """
        ...

    # region Properties

    @property
    def prop_width(self) -> UnitMM:
        """Gets or sets the width of the shape (in ``mm`` units)."""
        ...

    @prop_width.setter
    def prop_width(self, value: float | UnitT) -> None: ...

    @property
    def prop_height(self) -> UnitMM:
        """Gets or sets the height of the shape (in ``mm`` units)."""
        ...

    @prop_height.setter
    def prop_height(self, value: float | UnitT) -> None: ...

    # endregion Properties
