from __future__ import annotations
from typing import Any, Tuple
import uno
from com.sun.star.container import NoSuchElementException
from com.sun.star.container import ElementExistException
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.container import XNameContainer
from ooodev.uno_helper.base_class.base import Base


class NameContainerImpl(Base, XNameContainer):
    def __init__(self, element_type: Any):
        self._dict = {}
        self._element_type = element_type

    # region XElementAccess
    def getElementType(self) -> Any:
        """
        Returns the type of the elements.
        """
        return self._element_type

    def hasElements(self) -> bool:
        """
        Determines if the container has elements.
        """
        return bool(self._dict)

    # endregion XElementAccess

    # region XNameAccess
    def getByName(self, name: str) -> Any:
        """

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if name not in self._dict:
            raise NoSuchElementException(f"Element '{name}' not found", self)
        return self._dict[name]

    def getElementNames(self) -> Tuple[str, ...]:
        """
        The order of the names is not specified.
        """
        return tuple(self._dict.keys())

    def hasByName(self, name: str) -> bool:
        """
        In many cases the next call is XNameAccess.getByName(). You should optimize this case.
        """
        return name in self._dict

    # endregion XNameAccess

    # region XNameReplace
    def replaceByName(self, name: str, element: Any) -> None:
        """
        replaces the element with the specified name with the given element.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if name not in self._dict:
            raise NoSuchElementException(f"Element '{name}' not found", self)
        self._dict[name] = element

    # endregion XNameReplace

    # region XNameContainer
    def insertByName(self, name: str, element: Any) -> None:
        """
        inserts the given element at the specified name.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.ElementExistException: ``ElementExistException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if not name:
            raise IllegalArgumentException("Name cannot be empty", self, 0)
        if name in self._dict:
            raise ElementExistException(f"Element '{name}' already exists", self)
        try:
            self._dict[name] = element
        except Exception as e:
            raise IllegalArgumentException(f"Error inserting element '{name}'", self, 0)

    def removeByName(self, name: str) -> None:
        """
        removes the element with the specified name.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if name not in self._dict:
            raise NoSuchElementException(f"Element '{name}' not found", self)
        del self._dict[name]

    # endregion XNameContainer
