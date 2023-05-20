from __future__ import annotations
import contextlib
from typing import TypeVar, Type
import uno
from ooo.dyn.awt.size import Size as UnoSize
from ...proto.size_obj import SizeObj

_TSize = TypeVar(name="_TSize", bound="Size")


class Size:
    """
    Represents a size with positive values.

    See Also:
        :ref:`proto_size_obj`
    """

    def __init__(self, width: int, height: int) -> None:
        """
        Constructor

        Args:
            width (int): Width value.
            height (int): Height Value.
        """
        self._width = width
        self._height = height

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, Size):
            return self.width == oth.width and self.height == oth.height
        with contextlib.suppress(AttributeError):
            return self.width == oth.Width and self.height == oth.Height  # type: ignore
        return NotImplemented

    def swap(self) -> Size:
        """Gets an instance with values swapped."""
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
        self._width = value

    @property
    def height(self) -> int:
        """
        Gets/Sets height.
        """
        return self._height

    @height.setter
    def height(self, value: int):
        self._height = value

    @property
    def Width(self) -> int:
        return self._width

    @Width.setter
    def Width(self, value: int):
        self._width = value

    @property
    def Height(self) -> int:
        return self._height

    @Height.setter
    def Height(self, value: int):
        self._height = value
