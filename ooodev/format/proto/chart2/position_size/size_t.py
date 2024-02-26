from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

# pylint: disable=ungrouped-imports

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol

    from ooodev.units.unit_obj import UnitT
    from ooodev.units.unit_mm import UnitMM
else:
    Protocol = object
    UnitT = Any
    UnitMM = Any

# see ooodev.format.inner.direct.chart2.position_size.size.Size


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

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> SizeT:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            SizeT: New instance.
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> SizeT:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.
            **kwargs: Additional arguments.

        Returns:
            SizeT: New instance.
        """
        ...

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
