from __future__ import annotations
from typing import Any, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

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
    @override
    def addPropertyChangeListener(self, aPropertyName: str, xListener: XPropertyChangeListener) -> None:
        """
        Not Implemented.

        Adds an XPropertyChangeListener to the specified property.

        An empty name registers the listener to all bound properties. If the property is not bound, the behavior is not specified.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.

        Raises:
            NotImplementedError: ``NotImplementedError``
        """
        raise NotImplementedError

    @override
    def removePropertyChangeListener(self, aPropertyName: str, aListener: XPropertyChangeListener) -> None:
        """
        Not Implemented.

        Removes an XPropertyChangeListener from the listener list.

        It is a ``noop`` if the listener is not registered.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.

        Raises:
            NotImplementedError: ``NotImplementedError``
        """
        raise NotImplementedError

    @override
    def addVetoableChangeListener(self, PropertyName: str, aListener: XVetoableChangeListener) -> None:
        """
        Not Implemented.

        Adds an XVetoableChangeListener to the specified property with the name PropertyName.

        An empty name registers the listener to all constrained properties.
        If the property is not constrained, the behavior is not specified.

        Raises:
            NotImplementedError: ``NotImplementedError``
        """
        raise NotImplementedError

    @override
    def removeVetoableChangeListener(self, PropertyName: str, aListener: XVetoableChangeListener) -> None:
        """
        Not Implemented.

        Removes an XVetoableChangeListener from the listener list.

        It is a ``noop`` if the listener is not registered.

        Raises:
            NotImplementedError: ``NotImplementedError``
        """
        raise NotImplementedError

    @override
    def getPropertySetInfo(self) -> XPropertySetInfo:
        """
        Not Implemented.

        Gets the complete information of the properties provided by this object.

        Raises:
            NotImplementedError: ``NotImplementedError``
        """
        raise NotImplementedError

    @override
    def setPropertyValue(self, aPropertyName: str, aValue: Any) -> None:
        """
        Not Implemented.

        Sets the value of the property with the specified name.

        If it is a bound property the value will be changed before the change event is fired.
        If it is a constrained property a vetoable event is fired before the property value can be changed.

        Raises:
            NotImplementedError: ``NotImplementedError``
        """
        raise NotImplementedError

    @override
    def getPropertyValue(self, PropertyName: str) -> Any:
        """
        Not Implemented.

        Gets a property value.

        Raises:
            NotImplementedError: ``NotImplementedError``
        """
        raise NotImplementedError

    # endregion XPropertySet
