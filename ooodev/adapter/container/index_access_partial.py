from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.container import XIndexAccess

from .element_access_partial import ElementAccessPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class IndexAccessPartial(ElementAccessPartial):
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

    def __iter__(self):
        self.__current_index = -1
        return self

    def __next__(self):
        self.__current_index += 1
        if self.__current_index >= self.__component.getCount():
            self.__current_index = -1
            raise StopIteration
        # don't use self.get_by_index() here, it may be overridden in child class
        return self.__component.getByIndex(self.__current_index)

    def __len__(self) -> int:
        """.. versionadded:: 0.20.2"""
        return self.get_count()

    # region XIndexAccess
    def get_count(self) -> int:
        """
        Gets the number of elements contained in the container.

        Returns:
            int: The number of elements.
        """
        return self.__component.getCount()

    def get_by_index(self, idx: int) -> Any:
        """
        Gets the element at the specified index.

        Args:
            idx (int): The Zero-based index of the element.

        Returns:
            Any: The element at the specified index.
        """
        return self.__component.getByIndex(idx)

    # endregion XIndexAccess
