from __future__ import annotations
from typing import Any, Tuple, TYPE_CHECKING

from ooodev.uno_helper.base_class.base_property_set import BasePropertySet
from ooodev.uno_helper.base_class.base_service_info import BaseServiceInfo
from ooodev.io.log.named_logger import NamedLogger

if TYPE_CHECKING:
    from com.sun.star.beans import XPropertyChangeListener
    from com.sun.star.beans import XVetoableChangeListener
    from com.sun.star.beans import XPropertySetInfo
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
    def getImplementationName(self) -> str:
        """
        Provides the implementation name of the service implementation.
        """
        return "action_trigger_item"

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
        return ("com.sun.star.ui.ActionTrigger",)

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
        if name == "CommandURL":
            self._command_url = value
        elif name == "Text":
            self._text = value
        elif name == "HelpURL":
            self._help_url = value
        elif name == "SubContainer":
            self._sub_menu = value
        elif name == "Image":
            self._image = value
        else:
            raise AttributeError(f"Unknown property: {name}")

    def getPropertyValue(self, name: str) -> Any:
        """
        Gets a property value.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        if name == "CommandURL":
            return self._command_url
        elif name == "Text":
            return self._text
        elif name == "HelpURL":
            return self._help_url
        elif name == "SubContainer":
            return self._sub_menu
        elif name == "Image":
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
