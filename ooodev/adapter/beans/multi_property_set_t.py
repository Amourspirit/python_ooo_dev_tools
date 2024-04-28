from __future__ import annotations
from typing import Any, Tuple, Protocol
import uno

from com.sun.star.beans import XPropertiesChangeListener
from com.sun.star.beans import XPropertySetInfo


class MultiPropertySetT(Protocol):
    """
    Protocol class for XMultiPropertySet.
    """

    # region XMultiPropertySet

    def addPropertiesChangeListener(
        self,
        names: Tuple[str, ...],
        listener: XPropertiesChangeListener,
    ) -> None:
        """
        adds an XPropertiesChangeListener to the specified property with the specified names.

        The implementation can ignore the names of the properties and fire the event on all properties.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.

        Args:
            names (Tuple[str, ...]): One or more names of the properties to listen to.
            listener (XPropertiesChangeListener): The listener to be added.
        """
        ...

    def firePropertiesChangeEvent(self, names: Tuple[str, ...], listener: XPropertiesChangeListener) -> None:
        """
        Fires a sequence of PropertyChangeEvents to the specified listener.

        Args:
            names (Tuple[str, ...]): One or more names of the properties to fire.
            listener (XPropertiesChangeListener): The listener to be notified.
        """
        ...

    def getPropertySetInfo(self) -> XPropertySetInfo:
        """
        Gets the information of the properties which can be set at this object.
        """
        ...

    def getPropertyValues(self, names: Tuple[str, ...]) -> Tuple[Any, ...]:
        """
        The order of the values in the returned sequence will be the same as the order of the names in the argument.

        Args:
            names (Tuple[str, ...]): One or more names of the properties.
        """
        ...

    def removePropertiesChangeListener(self, listener: XPropertiesChangeListener) -> None:
        """
        removes an ``XPropertiesChangeListener`` from the listener list.

        It is a ``noop`` if the listener is not registered.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.
        """
        ...

    def setPropertyValues(self, names: Tuple[str, ...], values: Tuple[Any, ...]) -> None:
        """
        Sets the values to the properties with the specified names.

        The values of the properties must change before the bound events are fired. The values of the constrained properties should change after the vetoable events are fired and only if no exception occurred. Unknown properties are ignored.

        Similar to ``set_prop_values()`` but with tuples instead of key, value arguments.

        Raises:
            com.sun.star.beans.PropertyVetoException: ``PropertyVetoException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        ...

    # endregion XMultiPropertySet
