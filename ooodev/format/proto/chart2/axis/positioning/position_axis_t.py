from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
else:
    Protocol = object


class PositionAxisT(StyleT, Protocol):
    """Position Axis Protocol"""

    def __init__(self, on_mark: bool = ...) -> None:
        """
        Constructor

        Args:
            on_mark(bool, optional): Specifies that the axis is position.
                If ``True``, specifies that the axis is positioned on the first/last tickmarks. This makes the data points visual representation begin/end at the value axis.
                If ``False``, specifies that the axis is positioned between the tickmarks. This makes the data points visual representation begin/end at a distance from the value axis.

        Returns:
            None:
        """

        ...

    @property
    def prop_on_mark(self) -> bool: ...

    @prop_on_mark.setter
    def prop_on_mark(self, value: bool) -> None: ...
