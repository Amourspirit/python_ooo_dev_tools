from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING


from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_multi_t import StyleMultiT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Self
    from typing_extensions import Protocol
    from ooodev.units.unit_obj import UnitT
    from ooodev.format.inner.direct.structs.side import Side as Side
    from ooodev.format.inner.direct.structs.table_border_struct import TableBorderStruct
    from ooodev.format.inner.direct.calc.border.padding import Padding
    from ooodev.format.inner.direct.calc.border.shadow import Shadow

else:
    Protocol = object
    Self = Any
    UnitT = Any
    Side = Any
    TableBorderStruct = Any
    Padding = Any
    Shadow = Any


# see ooodev.format.inner.direct.calc.numbers.numbers.Numbers
class BordersT(StyleMultiT, Protocol):
    """Borders Protocol"""

    def __init__(
        self,
        *,
        right: Side | None = ...,
        left: Side | None = ...,
        top: Side | None = ...,
        bottom: Side | None = ...,
        border_side: Side | None = ...,
        vertical: Side | None = ...,
        horizontal: Side | None = ...,
        distance: float | UnitT | None = ...,
        diagonal_down: Side | None = ...,
        diagonal_up: Side | None = ...,
        shadow: Shadow | None = ...,
        padding: Padding | None = ...,
    ) -> None:
        """
        Constructor

        Args:
            left (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the left edge.
            right (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the right edge.
            top (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the top edge.
            bottom (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the bottom edge.
            border_side (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            horizontal (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style of horizontal lines for the inner part of a cell range.
            vertical (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style of vertical lines for the inner part of a cell range.
            distance (float, UnitT, optional): Contains the distance between the lines and other contents in ``mm`` units or :ref:`proto_unit_obj`.
            diagonal_down (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style from top-left to bottom-right diagonal.
            diagonal_up (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style from bottom-left to top-right diagonal.
            shadow (~ooodev.format.inner.direct.calc.border.shadow.Shadow, optional): Cell Shadow.
            padding (padding, optional): Cell padding.

        Returns:
            None:

        Hint:
            ``Side``, ``Shadow`` and ``Padding`` can be imported from ``ooodev.format.calc.direct.cell.borders``
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> BordersT:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            BordersT: Instance that represents numbers format.
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> BordersT:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.
            kwargs: Additional arguments.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            BordersT: Instance that represents numbers format.
        """
        ...

    # region Style Methods
    def fmt_border_side(self, value: Side | None) -> BordersT:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        ...

    def fmt_left(self, value: Side | None) -> BordersT:
        """
        Gets copy of instance with left set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        ...

    def fmt_right(self, value: Side | None) -> BordersT:
        """
        Gets copy of instance with right set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        ...

    def fmt_top(self, value: Side | None) -> BordersT:
        """
        Gets copy of instance with top set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        ...

    def fmt_bottom(self, value: Side | None) -> BordersT:
        """
        Gets copy of instance with bottom set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        ...

    def fmt_horizontal(self, value: Side | None) -> BordersT:
        """
        Gets copy of instance with horizontal set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        ...

    def fmt_vertical(self, value: Side | None) -> BordersT:
        """
        Gets copy of instance with vertical set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        ...

    def fmt_distance(self, value: float | UnitT | None) -> BordersT:
        """
        Gets copy of instance with distance set or removed

        Args:
            value (float, UnitT, optional): Distance value

        Returns:
            Borders: Borders instance
        """
        ...

    def fmt_diagonal_down(self, value: Side | None) -> BordersT:
        """
        Gets copy of instance with diagonal down set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        ...

    def fmt_diagonal_up(self, value: Side | None) -> BordersT:
        """
        Gets copy of instance with diagonal up set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        ...

    def fmt_shadow(self, value: Shadow | None) -> BordersT:
        """
        Gets copy of instance with shadow set or removed

        Args:
            value (Shadow, optional): Shadow value

        Returns:
            Borders: Borders instance
        """
        ...

    def fmt_padding(self, value: Padding | None) -> BordersT:
        """
        Gets copy of instance with padding set or removed

        Args:
            value (Padding, optional): Padding value

        Returns:
            Borders: Borders instance
        """
        ...

    # endregion Style Methods

    # region Properties

    @property
    def prop_inner_padding(self) -> Padding:
        """Gets Padding instance"""
        ...

    @property
    def prop_inner_border_table(self) -> TableBorderStruct:
        """Gets border table instance"""
        ...

    @property
    def prop_inner_shadow(self) -> Shadow | None:
        """Gets inner shadow instance"""
        ...

    @property
    def prop_inner_diagonal_up(self) -> Side | None:
        """Gets inner Diagonal up instance"""
        ...

    @property
    def prop_inner_diagonal_dn(self) -> Side | None:
        """Gets inner Diagonal down instance"""
        ...

    @property
    def default(self) -> BordersT:  # type: ignore[misc]
        """Gets Default Border."""
        ...

    @property
    def empty(self) -> BordersT:  # type: ignore[misc]
        """Gets Empty Border. When style is applied formatting is removed."""
        ...

    # endregion Properties
