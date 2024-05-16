from __future__ import annotations
from typing import Any, Tuple, Sequence, overload
import contextlib

import uno

# import unohelper
from com.sun.star.container import XElementAccess
from com.sun.star.container import XIndexAccess
from com.sun.star.uno import XInterface
from com.sun.star.lang import XTypeProvider


# see also https://ask.libreoffice.org/t/create-python-xinterface-implementation/104666


class IndexAccessImplement(XTypeProvider, XIndexAccess, XElementAccess, XInterface):
    """
    Index Access that implements ``XIndexAccess``

    Note:
        Supports iteration including reversed, slicing, and indexing.
    """

    __pyunointerface__: str = "com.sun.star.container.XIndexAccess"

    def __init__(self, elements: Sequence[Any], element_type: str = "[]any"):
        """
        Initializes the Index Access

        Args:
            elements (Sequence[Any]): Any sequence of elements that can be accessed by index such as a tuple
            element_type (str, optional): Used to get the Element type of this instance. Defaults to "[]any".

        Return:
            None:

        Note:
            ``element_type`` is used to set the Element type of this instance.
            Can be value such as:

            - ``[][]com.sun.star.beans.PropertyValue``, a tuple of tuple of PropertyValue
            - ``[]com.sun.star.beans.PropertyValue``, a tuple of PropertyValue
            - ``[][]long``, a tuple of tuple of int
            - ``[]long``, a tuple of int
            - ``[][]string``, a tuple of tuple of str
            - ``[]string``, a tuple of str
            - ``[]any``, a tuple of Any.

            List are also supported but depending on how the instance is used, it may have unexpected results.

        See Also:
            `UNO Language binding <https://www.openoffice.org/udk/python/python-bridge.html#binding>`__
        """
        super().__init__()
        self._data = elements
        self._types = None
        self._element_type = element_type
        self._iter_index = 0

    # region Dunder Methods
    def __len__(self) -> int:
        return self.getCount()

    def __iter__(self):
        self._iter_index = 0
        length = len(self)
        while self._iter_index < length:
            yield self._data[self._iter_index]
            self._iter_index += 1

    def __reversed__(self):
        self._iter_index = len(self) - 1
        while self._iter_index >= 0:
            yield self._data[self._iter_index]
            self._iter_index -= 1

    # region __getitem__
    @overload
    def __getitem__(self, index: int) -> Any: ...

    @overload
    def __getitem__(self, index: slice) -> IndexAccessImplement: ...

    def __getitem__(self, index: int | slice) -> Any:
        """
        Get an item or a slice of items from this instance.

        Args:
            index (int or slice): The index or slice to get.

        Returns:
            The item or slice of items at the given index or slice.
        """
        if isinstance(index, slice):
            return type(self)(self._data[index], self._element_type)
        else:
            return self._data[index]

    # endregion __getitem__

    # endregion Dunder Methods

    # region XInterface
    def acquire(self) -> None:
        pass

    def release(self) -> None:
        pass

    def queryInterface(self, a_type: Any) -> Any:
        with contextlib.suppress(Exception):
            if a_type in self.getTypes():
                return self
        return None

    # end region XInterface

    # region XIndexAccess
    def getCount(self) -> int:
        """
        Gets the number of elements in the collection.

        Returns:
            int: number of elements in the collection.
        """
        return len(self._data)

    def getByIndex(self, index: int) -> Any:
        """
        Gets the element at the specified index.

        Args:
            index (int): index of element to return.

        Returns:
            Any: element at the specified index.
        """
        return self._data[index]

    # endregion XIndexAccess

    # region XElementAccess
    def hasElements(self) -> bool:
        """
        Determines whether the collection has elements.

        Returns:
            bool: ``True`` if the collection has elements, otherwise ``False``.
        """
        return self.getCount() > 0

    def getElementType(self) -> Any:
        """
        Gets te Element Type
        """
        t = uno.getTypeByName(self._element_type)
        return t

    # endregion XElementAccess

    # region XTypeProvider
    def getImplementationId(self) -> Any:
        """
        Obsolete unique identifier.
        """
        return b""

    def getTypes(self) -> Tuple[Any, ...]:
        """
        returns a sequence of all types (usually interface types) provided by the object.
        """
        if self._types is None:
            types = []
            types.append(uno.getTypeByName("com.sun.star.uno.XInterface"))
            types.append(uno.getTypeByName("com.sun.star.lang.XTypeProvider"))
            types.append(uno.getTypeByName("com.sun.star.container.XElementAccess"))
            types.append(uno.getTypeByName("com.sun.star.container.XIndexAccess"))
            self._types = tuple(types)
        return self._types

    # endregion XTypeProvider
