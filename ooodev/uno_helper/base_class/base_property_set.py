from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.beans import XPropertySet
from ooodev.uno_helper.base_class.base import Base

if TYPE_CHECKING:
    from com.sun.star.beans import XPropertyChangeListener
    from com.sun.star.beans import XVetoableChangeListener
    from com.sun.star.beans import XPropertySetInfo


class BasePropertySet(Base, XPropertySet):
    """
    Provides information about and access to the properties from an implementation.

    There are three types of properties:

    You can listen to changes of bound properties with the XPropertyChangeListener and you can veto changes of constrained properties with the XVetoableChangeListener.

    To implement inaccurate name access, you must support the interface XExactName.

    See Also:
        `API XPropertySet <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertySet.html>`_
    """

    # region XPropertySet
    def addPropertyChangeListener(self, listener: XPropertyChangeListener, prop_name: str = "") -> None:
        """
        Adds an XPropertyChangeListener to the specified property.

        An empty name registers the listener to all bound properties. If the property is not bound, the behavior is not specified.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        raise NotImplementedError

    def removePropertyChangeListener(self, listener: XPropertyChangeListener, prop_name: str = "") -> None:
        """
        removes an XPropertyChangeListener from the listener list.

        It is a ``noop`` if the listener is not registered.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        raise NotImplementedError

    def addVetoableChangeListener(self, listener: XVetoableChangeListener, prop_name: str = "") -> None:
        """
        Adds an XVetoableChangeListener to the specified property with the name PropertyName.

        An empty name registers the listener to all constrained properties.
        If the property is not constrained, the behavior is not specified.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        raise NotImplementedError

    def removeVetoableChangeListener(self, listener: XVetoableChangeListener, prop_name: str = "") -> None:
        """
        removes an XVetoableChangeListener from the listener list.

        It is a ``noop`` if the listener is not registered.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        raise NotImplementedError

    def getPropertySetInfo(self) -> XPropertySetInfo:
        """
        Gets the complete information of the properties provided by this object.

        Returns:
            XPropertySetInfo: Property set info.
        """
        raise NotImplementedError

    def setPropertyValue(self, name: str, value: Any) -> None:
        """
        Sets the value of the property with the specified name.

        If it is a bound property the value will be changed before the change event is fired.
        If it is a constrained property a vetoable event is fired before the property value can be changed.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.beans.PropertyVetoException: ``PropertyVetoException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        raise NotImplementedError

    def getPropertyValue(self, name: str) -> Any:
        """
        Gets a property value.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        raise NotImplementedError

    # endregion XPropertySet
