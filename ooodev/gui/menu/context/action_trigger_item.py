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
    from com.sun.star.container import XIndexContainer
    from com.sun.star.awt import XBitmap


class ActionTriggerItem(BasePropertySet, BaseServiceInfo):
    """Provides item."""

    def __init__(
        self,
        command_url: str,
        text: str,
        help_url: str = "",
        sub_menu: XIndexContainer | None = None,
        image: XBitmap | None = None,
    ):
        """
        Constructor.

        Args:
            command_url (str): Contains the command URL for the menu entry.
            text (str): Contains the text of the menu entry.
            help_url (str): Contains the a URL that points to a help text.
            sub_menu (XIndexContainer | None, optional): Contains a sub menu. Defaults to None.
        """
        self._command_url = command_url
        self._text = text
        self._help_url = help_url
        self._sub_menu = sub_menu
        self._image = image
        self._logger = NamedLogger(self.__class__.__name__)

    # region XServiceInfo
    @override
    def getImplementationName(self) -> str:
        """
        Provides the implementation name of the service implementation.
        """
        return "action_trigger_item"

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
        return ("com.sun.star.ui.ActionTrigger",)

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
        if aPropertyName == "CommandURL":
            self._command_url = aValue
        elif aPropertyName == "Text":
            self._text = aValue
        elif aPropertyName == "HelpURL":
            self._help_url = aValue
        elif aPropertyName == "SubContainer":
            self._sub_menu = aValue
        elif aPropertyName == "Image":
            self._image = aValue
        else:
            raise AttributeError(f"Unknown property: {aPropertyName}")

    @override
    def getPropertyValue(self, PropertyName: str) -> Any:
        """
        Gets a property value.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if PropertyName == "CommandURL":
            return self._command_url
        elif PropertyName == "Text":
            return self._text
        elif PropertyName == "HelpURL":
            return self._help_url
        elif PropertyName == "SubContainer":
            return self._sub_menu
        elif PropertyName == "Image":
            return self._image
        else:
            return None

    # endregion XPropertySet

    # region Other methods
    def is_separator(self) -> bool:
        """
        Gets if the item is a separator.

        Returns:
            bool: ``False``
        """
        return False

    # endregion Other methods

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__}({self._command_url}, {self._text}, {self._help_url}, {self._sub_menu}, {self._image})>"

    # region Properties
    @property
    def CommandURL(self) -> str:
        """Command URL"""
        return self._command_url

    @CommandURL.setter
    def CommandURL(self, value: str) -> None:
        self._command_url = value

    @property
    def Text(self) -> str:
        """Text"""
        return self._text

    @Text.setter
    def Text(self, value: str) -> None:
        self._text = value

    @property
    def HelpURL(self) -> str:
        """Help URL"""
        return self._help_url

    @HelpURL.setter
    def HelpURL(self, value: str) -> None:
        self._help_url = value

    @property
    def SubContainer(self) -> XIndexContainer | None:
        """Sub Container"""
        return self._sub_menu

    @SubContainer.setter
    def SubContainer(self, value: XIndexContainer | None) -> None:
        self._sub_menu = value

    @property
    def Image(self) -> XBitmap | None:
        """Image"""
        return self._image

    @Image.setter
    def Image(self, value: XBitmap | None) -> None:
        self._image = value

    # endregion Properties
