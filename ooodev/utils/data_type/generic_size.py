from __future__ import annotations
import contextlib
from typing import TypeVar, Generic, Union, TYPE_CHECKING
import uno
from ooo.dyn.awt.size import Size as UnoSize

T = TypeVar(name="T", bound=Union[int, float])

if TYPE_CHECKING:
    from typing_extensions import Self


class GenericSize(Generic[T]):
    """
    Represents a generic size.

    See Also:
        :ref:`proto_size_obj`, :py::class:`ooodev.utils.data_type.size.Size`

    .. versionadded:: 0.14.0
    """

    def __init__(self, width: T, height: T) -> None:
        """
        Constructor

        Args:
            width (int): Width value.
            height (int): Height Value.
        """
        self._width = width
        self._height = height

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, GenericSize):
            return self.width == oth.width and self.height == oth.height
        with contextlib.suppress(AttributeError):
            return self.width == oth.Width and self.height == oth.Height  # type: ignore
        return NotImplemented

    def swap(self) -> Self:
        """Gets an instance with values swapped."""
        return self.__class__(self.height, self.width)

    def get_uno_size(self) -> UnoSize:
        """Gets UNO instance from current values"""
        return UnoSize(int(self.width), int(self.height))

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
