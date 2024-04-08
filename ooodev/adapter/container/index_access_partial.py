from __future__ import annotations
from typing import Any, TYPE_CHECKING, Generic, TypeVar
import uno

from com.sun.star.container import XIndexAccess

from ooodev.adapter.container.element_access_partial import ElementAccessPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface

T = TypeVar("T")


class IndexAccessPartial(Generic[T], ElementAccessPartial):
    """
    Partial Class for XIndexAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XIndexAccess, interface: UnoInterface | None = XIndexAccess) -> None:
        """
        Constructor

        Args:
            component (XIndexAccess): UNO Component that implements ``com.sun.star.container.XIndexAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XIndexAccess``.
        """
        ElementAccessPartial.__init__(self, component, interface)
        self.__component = component
        self.__current_index = -1

    def _is_next_index_element_valid(self, element: Any) -> bool:
        """
        Gets if the next element is valid.
        This method is called when iterating over the elements of this class.

        Args:
            element (Any): Element

        Returns:
            bool: True in this class but can be overridden in child classes.
        """
        return True

    def __iter__(self):
        self.__current_index = -1
        return self

    def __next__(self) -> T:
        self.__current_index += 1
        if self.__current_index >= self.__component.getCount():
            self.__current_index = -1
            raise StopIteration
        # don't use self.get_by_index() here, it may be overridden in child class
        next_element = self.__component.getByIndex(self.__current_index)
        if self._is_next_index_element_valid(next_element):
            return next_element
        # this method may be overridden in child classes and called with super()
        # Call recursively using IndexAccessPartial
        return IndexAccessPartial.__next__(self)

    def __len__(self) -> int:
        """.. versionadded:: 0.20.2"""
        return self.get_count()

    def __getitem__(self, idx: int) -> T:
        """Get By Index"""
        return self.get_by_index(idx)

    # region XIndexAccess
    def get_count(self) -> int:
        """
        Gets the number of elements contained in the container.

        Returns:
            int: The number of elements.
        """
        return self.__component.getCount()

    def get_by_index(self, idx: int) -> T:
        """
        Gets the element at the specified index.

        Args:
            idx (int): The Zero-based index of the element.

        Returns:
            Any: The element at the specified index.
        """
        return self.__component.getByIndex(idx)

    # endregion XIndexAccess
