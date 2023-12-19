from __future__ import annotations
import contextlib
from typing import TypeVar, Generic, Union, TYPE_CHECKING
import uno
from ooo.dyn.awt.point import Point as UnoPoint

T = TypeVar(name="T", bound=Union[int, float])

if TYPE_CHECKING:
    from typing_extensions import Self


class GenericPoint(Generic[T]):
    """
    Represents a generic point.

    See Also:
        :ref:`proto_size_obj`, :py::class:`ooodev.utils.data_type.size.Size`

    .. versionadded:: 0.17.3
    """

    def __init__(self, x: T, y: T) -> None:
        """
        Constructor

        Args:
            x (int): X value.
            y (int): Y Value.
        """
        self._x = x
        self._y = y

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, GenericPoint):
            return self.x == oth.x and self.y == oth.y
        with contextlib.suppress(AttributeError):
            return self.x == oth.x and self.y == oth.y  # type: ignore
        with contextlib.suppress(AttributeError):
            return self.x == oth.X and self.y == oth.Y  # type: ignore
        return NotImplemented

    def swap(self) -> Self:
        """Gets an instance with values swapped."""
        return self.__class__(self.y, self.x)

    def get_uno_point(self) -> UnoPoint:
        """Gets UNO instance from current values"""
        return UnoPoint(int(self.x), int(self.y))

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
