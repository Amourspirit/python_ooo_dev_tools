from __future__ import annotations
from typing import Generic, TypeVar
import uno

from com.sun.star.container import XIndexContainer

from ooodev.utils.type_var import UnoInterface
from ooodev.adapter.container.index_replace_partial import IndexReplacePartial

T = TypeVar("T")


class IndexContainerPartial(IndexReplacePartial[T], Generic[T]):
    """
    Partial Class for XIndexContainer.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XIndexContainer, interface: UnoInterface | None = XIndexContainer) -> None:
        """
        Constructor

        Args:
            component (XIndexContainer): UNO Component that implements ``com.sun.star.container.XIndexContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XIndexContainer``.
        """
        IndexReplacePartial.__init__(self, component, interface)
        self.__component = component

    # region XIndexContainer
    def insert_by_index(self, index: int, element: T) -> None:
        """
        Inserts the given element at the specified index.

        To append an element, use the index ``last index +1``.

        Args:
            index (int): The Zero-based index at which the element should be inserted.
            element (T): The element to insert.

        Raises:
            IllegalArgumentException: ``com.sun.star.lang.IllegalArgumentException``
            IndexOutOfBoundsException: ``com.sun.star.lang.IndexOutOfBoundsException``
            WrappedTargetException: ``com.sun.star.lang.WrappedTargetException``
        """
        self.__component.insertByIndex(index, element)

    def remove_by_index(self, index: int) -> None:
        """
        Removes the element at the specified index.

        Args:
            index (int): The Zero-based index of the element to remove.

        Raises:
            IndexOutOfBoundsException: ``com.sun.star.lang.IndexOutOfBoundsException``
            WrappedTargetException: ``com.sun.star.lang.WrappedTargetException``
        """
        self.__component.removeByIndex(index)

    # endregion XIndexContainer
