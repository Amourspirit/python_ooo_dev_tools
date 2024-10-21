from __future__ import annotations
from typing import Any, Tuple, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.uno_helper.base_class.base_property_set import BasePropertySet
from ooodev.uno_helper.base_class.base_service_info import BaseServiceInfo
from ooodev.io.log.named_logger import NamedLogger

if TYPE_CHECKING:
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
    @override
    def getImplementationName(self) -> str:
        """
        Provides the implementation name of the service implementation.
        """
        return "action_trigger_sep"

    @override
    def supportsService(self, ServiceName: str) -> bool:
        """
        Tests whether the specified service is supported, i.e.

        implemented by the implementation.
        """
        return ServiceName in self.getSupportedServiceNames()

    @override
    def getSupportedServiceNames(self) -> Tuple[str]:
        """
        Provides the supported service names of the implementation, including also indirect service names.
        """
        return ("com.sun.star.ui.ActionTriggerSeparator",)

    # endregion XServiceInfo

    # region XPropertySet

    @override
    def setPropertyValue(self, aPropertyName: str, aValue: Any) -> None:
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
        if aPropertyName == "SeparatorType":
            self.separator_type = aValue
        else:
            raise AttributeError(f"Unknown property: {aPropertyName}")

    def getPropertyValue(self, PropertyName: str) -> Any:
        """
        Gets a property value.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if PropertyName == "SeparatorType":
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
