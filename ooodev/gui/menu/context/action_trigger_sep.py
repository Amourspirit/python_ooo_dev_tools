from __future__ import annotations
from typing import Any, Tuple, TYPE_CHECKING
import uno

from ooodev.uno_helper.base_class.base_property_set import BasePropertySet
from ooodev.uno_helper.base_class.base_service_info import BaseServiceInfo
from ooodev.io.log.named_logger import NamedLogger

if TYPE_CHECKING:
    from com.sun.star.beans import XPropertyChangeListener
    from com.sun.star.beans import XVetoableChangeListener
    from com.sun.star.beans import XPropertySetInfo
    from ooo.dyn.ui.action_trigger_separator_type import ActionTriggerSeparatorTypeEnum


class ActionTriggerSep(BasePropertySet, BaseServiceInfo):
    """Provides item."""

    def __init__(
        self,
        separator_type: ActionTriggerSeparatorTypeEnum | int = 0,
    ):
        """
        Constructor.

        Args:
            separator_type (int, ActionTriggerSeparatorTypeEnum, optional): Separator Type. Defaults to ``0`` (Line).

        Hint:
            - ``ActionTriggerSeparatorTypeEnum`` is an enum and can be imported from ``ooo.dyn.ui.action_trigger_separator_type``.

            ``separator_type`` can be one of the following:

            - ``0``: Line
            - ``1``: Space
            - ``2``: LineBreak
        """
        self._separator_type = int(separator_type)
        self._logger = NamedLogger(self.__class__.__name__)

    # region XServiceInfo
    def getImplementationName(self) -> str:
        """
        Provides the implementation name of the service implementation.
        """
        return "action_trigger_sep"

    def supportsService(self, name: str) -> bool:
        """
        Tests whether the specified service is supported, i.e.

        implemented by the implementation.
        """
        return name in self.getSupportedServiceNames()

    def getSupportedServiceNames(self) -> Tuple[str]:
        """
        Provides the supported service names of the implementation, including also indirect service names.
        """
        return ("com.sun.star.ui.ActionTriggerSeparator",)

    # endregion XServiceInfo

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
        self._logger.warning("addPropertyChangeListener is not implemented in this context")
        return None  # type: ignore

    def removePropertyChangeListener(self, listener: XPropertyChangeListener, prop_name: str = "") -> None:
        """
        removes an XPropertyChangeListener from the listener list.

        It is a ``noop`` if the listener is not registered.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self._logger.warning("removePropertyChangeListener is not implemented in this context")
        return None  # type: ignore

    def addVetoableChangeListener(self, listener: XVetoableChangeListener, prop_name: str = "") -> None:
        """
        Adds an XVetoableChangeListener to the specified property with the name PropertyName.

        An empty name registers the listener to all constrained properties.
        If the property is not constrained, the behavior is not specified.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self._logger.warning("addVetoableChangeListener is not implemented in this context")
        return None  # type: ignore

    def removeVetoableChangeListener(self, listener: XVetoableChangeListener, prop_name: str = "") -> None:
        """
        removes an XVetoableChangeListener from the listener list.

        It is a ``noop`` if the listener is not registered.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self._logger.warning("removeVetoableChangeListener is not implemented in this context")
        return None  # type: ignore

    def getPropertySetInfo(self) -> XPropertySetInfo:
        """
        Gets the complete information of the properties provided by this object.

        Returns:
            XPropertySetInfo: Property set info.
        """
        self._logger.warning("getPropertySetInfo is not implemented in this context")
        return None  # type: ignore

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
        if name == "SeparatorType":
            self.separator_type = value
        else:
            raise AttributeError(f"Unknown property: {name}")

    def getPropertyValue(self, name: str) -> Any:
        """
        Gets a property value.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if name == "SeparatorType":
            return self.separator_type
        else:
            return None

    # endregion XPropertySet

    # region Other methods
    def is_separator(self) -> bool:
        """
        Gets if the item is a separator.

        Returns:
            bool: ``True``
        """
        return True

    # endregion Other methods

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(SeparatorType={self.SeparatorType})"

    # region Properties
    @property
    def SeparatorType(self) -> int:
        """Command URL"""
        return self._separator_type

    @SeparatorType.setter
    def SeparatorType(self, value: int | ActionTriggerSeparatorTypeEnum) -> None:
        self._separator_type = int(value)

    # endregion Properties
