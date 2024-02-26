from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooodev.utils.color import Color
else:
    Protocol = object
    Color = Any


class AbstractFillColor(StyleT, Protocol):
    """Abstract Fill Protocol"""

    def __init__(self, *, color: Color = ...) -> None: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> AbstractFillColor: ...
    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> AbstractFillColor: ...
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> AbstractFillColor: ...

    @property
    def prop_color(self) -> Color:
        """Gets/Sets color"""
        ...

    @prop_color.setter
    def prop_color(self, value: Color): ...

    @property
    def default(self) -> AbstractFillColor:
        """Gets Color empty."""
        ...
