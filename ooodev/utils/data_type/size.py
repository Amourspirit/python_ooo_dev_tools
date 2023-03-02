from __future__ import annotations
from typing import TypeVar, Type
import uno
from ooo.dyn.awt.size import Size as UnoSize
from ...proto.size_obj import SizeObj

_TSize = TypeVar(name="_TSize", bound="Size")


class Size:
    """Represents a size with postive values."""

    def __init__(self, width: int, height: int) -> None:
        """
        Constructor

        Args:
            width (int): Width value.
            height (int): Height Value.
        """
        self.width = width
        self.height = height

    def swap(self) -> Size:
        """Gets an instance with values swaped."""
        return Size(self.height, self.width)

    def get_uno_size(self) -> UnoSize:
        """Gets UNO instance from current values"""
        return UnoSize(self.width, self.height)

    @classmethod
    def from_size(cls: Type[_TSize], sz: SizeObj) -> _TSize:
        """
        Gets instance from Size.

        Args:
            sz (Size): Size object, Can be UNO Size.

        Returns:
            Size: Size instance from Size values.
        """
        inst = super(Size, cls).__new__(cls)
        inst.__init__(sz.Width, sz.Height)
        return inst

    @property
    def width(self) -> int:
        """
        Gets/Sets width.
        """
        return self._width

    @width.setter
    def width(self, value: int):
        self._width = int(value)

    @property
    def height(self) -> int:
        """
        Gets/Sets height.
        """
        return self._height

    @height.setter
    def height(self, value: int):
        self._height = int(value)

    Width = width
    Height = height
