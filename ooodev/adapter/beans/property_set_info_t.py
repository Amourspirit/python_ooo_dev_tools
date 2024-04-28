from __future__ import annotations
from typing import Any, Protocol
import uno
from com.sun.star.beans import XPropertyChangeListener
from com.sun.star.beans import XVetoableChangeListener
from com.sun.star.beans import XPropertySetInfo


class PropertySetInfoT(Protocol):
    """Protocol class for XPropertySetInfo."""

    # region XPropertySet
    def addPropertyChangeListener(self, name: str, listener: XPropertyChangeListener) -> None:
        """
        Adds a listener for property changes.

        Args:
            name (str): The name of the property.
            listener (Any): The listener to be added.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        ...

    def addVetoableChangeListener(self, name: str, listener: XVetoableChangeListener) -> None:
        """
        Adds a listener for vetoable changes.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        ...

    def getPropertySetInfo(self) -> XPropertySetInfo:
        """
        Returns the property set info.

        Returns:
            XPropertySetInfo: The property set info.
        """
        ...

    def getPropertyValue(self, name: str) -> Any:
        """
        Returns the value of a property.

        Args:
            name (str): The name of the property.

        Returns:
            Any: The value of the property.
        """
        ...

    def removePropertyChangeListener(self, name: str, listener: XPropertyChangeListener) -> None:
        """
        Removes a listener for property changes.

        Args:
            name (str): The name of the property.
            listener (Any): The listener to be removed.
        """
        ...

    def removeVetoableChangeListener(self, name: str, listener: XVetoableChangeListener) -> None:
        """
        Removes a listener for vetoable changes.

        Args:
            name (str): The name of the property.
            listener (Any): The listener to be removed.
        """
        ...

    def setPropertyValue(self, name: str, value: Any) -> None:
        """
        Sets the value of a property.

        Args:
            name (str): The name of the property.
            value (Any): The value of the property.
        """
        ...

    # endregion XPropertySet
