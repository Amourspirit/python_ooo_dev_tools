from __future__ import annotations
from typing import Any, Protocol


class FastPropertySetT(Protocol):
    """
    Protocol class for XFastPropertySet.
    """

    # region XFastPropertySet
    def getFastPropertyValue(self, handle: int) -> Any:
        """
        returns the value of the property with the specified name.

        Args:
            handle (int): The implementation handle of the implementation for the property.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        ...

    def setFastPropertyValue(self, handle: int, value: Any) -> None:
        """
        Sets the value to the property with the specified name.

        Args:
            handle (int): The implementation handle of the implementation for the property.
            value (Any): The new value for the property.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.beans.PropertyVetoException: ``PropertyVetoException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        ...

    # endregion XFastPropertySet
