from __future__ import annotations
from typing import Any, Tuple
from com.sun.star.container import NoSuchElementException
from com.sun.star.container import ElementExistException
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.container import XNameContainer
from ooodev.uno_helper.base_class.base import Base


class NameContainerImpl(Base, XNameContainer):
    """
    Class that implements XNameContainer Component.
    """

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
    def getByName(self, aName: str) -> Any:
        """

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if aName not in self._dict:
            raise NoSuchElementException(f"Element '{aName}' not found", self)
        return self._dict[aName]

    def getElementNames(self) -> Tuple[str, ...]:
        """
        The order of the names is not specified.
        """
        return tuple(self._dict.keys())

    def hasByName(self, aName: str) -> bool:
        """
        In many cases the next call is XNameAccess.getByName(). You should optimize this case.
        """
        return aName in self._dict

    # endregion XNameAccess

    # region XNameReplace
    def replaceByName(self, aName: str, aElement: Any) -> None:
        """
        replaces the element with the specified name with the given element.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if aName not in self._dict:
            raise NoSuchElementException(f"Element '{aName}' not found", self)
        self._dict[aName] = aElement

    # endregion XNameReplace

    # region XNameContainer
    def insertByName(self, aName: str, aElement: Any) -> None:
        """
        inserts the given element at the specified name.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.ElementExistException: ``ElementExistException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if not aName:
            raise IllegalArgumentException("Name cannot be empty", self, 0)
        if aName in self._dict:
            raise ElementExistException(f"Element '{aName}' already exists", self)
        try:
            self._dict[aName] = aElement
        except Exception:
            raise IllegalArgumentException(f"Error inserting element '{aName}'", self, 0)

    def removeByName(self, Name: str) -> None:
        """
        removes the element with the specified name.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if Name not in self._dict:
            raise NoSuchElementException(f"Element '{Name}' not found", self)
        del self._dict[Name]

    # endregion XNameContainer
