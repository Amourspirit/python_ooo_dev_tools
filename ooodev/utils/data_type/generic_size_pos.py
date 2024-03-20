from __future__ import annotations
import contextlib
from typing import TypeVar, Generic, Union
import uno
from ooo.dyn.awt.rectangle import Rectangle as UnoRectangle

T = TypeVar(name="T", bound=Union[int, float])


class GenericSizePos(Generic[T]):
    """
    Represents a generic size and Position

    See Also:
        :ref:`proto_size_obj`, :py::class:`ooodev.utils.data_type.size.Size`

    .. versionadded:: 0.14.0
    """

    def __init__(self, x: T, y: T, width: T, height: T) -> None:
        """
        Constructor

        Args:
            x (T): X value.
            y (T): Y value.
            width (T): Width value.
            height (T): Height Value.
        """
        self._x = x
        self._y = y
        self._width = width
        self._height = height

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, GenericSizePos):
            return (
                int(self.width) == int(oth.width)
                and int(self.height) == int(oth.height)
                and int(self.x) == int(oth.x)
                and int(self.y) == int(oth.y)
            )
        with contextlib.suppress(AttributeError):
            return self.width == oth.Width and self.height == oth.Height  # type: ignore
        return NotImplemented

    def get_uno_size(self) -> UnoRectangle:
        """Gets UNO instance from current values"""
        return UnoRectangle(int(self.x), int(self.y), int(self.width), int(self.height))

    @property
    def x(self) -> T:
        """
        Gets/Sets x.
        """
        return self._x

    @x.setter
    def x(self, value: T):
        self._x = value

    @property
    def y(self) -> T:
        """
        Gets/Sets y.
        """
        return self._y

    @y.setter
    def y(self, value: T):
        self._y = value

    @property
    def width(self) -> T:
        """
        Gets/Sets width.
        """
        return self._width

    @width.setter
    def width(self, value: T):
        self._width = value

    @property
    def height(self) -> T:
        """
        Gets/Sets height.
        """
        return self._height

    @height.setter
    def height(self, value: T):
        self._height = value

    @property
    def Width(self) -> T:
        return self._width

    @Width.setter
    def Width(self, value: T):
        self._width = value

    @property
    def Height(self) -> T:
        return self._height

    @Height.setter
    def Height(self, value: T):
        self._height = value
