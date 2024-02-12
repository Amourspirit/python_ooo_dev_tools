from __future__ import annotations
from typing import Any, TYPE_CHECKING


if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
    from ooodev.format.proto.style_t import StyleT
else:
    Protocol = object
    StyleT = Any


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
