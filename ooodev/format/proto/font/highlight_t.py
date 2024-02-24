from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from ooodev.utils.color import Color
    from typing_extensions import Protocol
else:
    Protocol = object
    Color = Any


class HighlightT(StyleT, Protocol):
    """Font Highlight Protocol"""

    def __init__(self, color: Color = ...) -> None:
        """
        Constructor

        Args:
            color (~ooodev.utils.color.Color, optional): Highlight Color. A value of ``-1`` Set color to Transparent.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> HighlightT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> HighlightT: ...

    # region Format Methods
    def fmt_color(self, value: Color) -> HighlightT:
        """
        Gets copy of instance with color set.

        Args:
            value (~ooodev.utils.color.Color): color value. If value is less than zero it means no color.

        Returns:
            HighlightT: Highlight instance
        """
        ...

    # endregion Format Methods

    # region Properties

    @property
    def prop_color(self) -> Color:
        """Gets/Sets color"""
        ...

    @prop_color.setter
    def prop_color(self, value: Color): ...

    @property
    def empty(self) -> HighlightT:  # type: ignore[misc]
        """Gets Highlight empty."""
        ...

    # endregion Properties
