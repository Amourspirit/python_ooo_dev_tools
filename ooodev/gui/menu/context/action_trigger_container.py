from __future__ import annotations
from typing import Any, Tuple
import uno
from com.sun.star.container import XIndexContainer
from com.sun.star.lang import XMultiServiceFactory
from ooodev.io.log.named_logger import NamedLogger
from ooodev.uno_helper.base_class.base import Base
from ooodev.uno_helper.base_class.base_service_info import BaseServiceInfo


class ActionTriggerContainer(Base, BaseServiceInfo, XIndexContainer, XMultiServiceFactory):
    """Provides container for items."""

    # region Init
    def __init__(self):
        self._items = []
        self._logger = NamedLogger(self.__class__.__name__)

    # endregion Init

    # region BaseServiceInfo
    def getImplementationName(self) -> str:
        """
        Provides the implementation name of the service implementation.
        """
        return "action_trigger_container"

    def supportsService(self, name: str) -> bool:
        """
        Tests whether the specified service is supported, i.e.

        implemented by the implementation.
        """
        return name in self.getAvailableServiceNames()

    def getSupportedServiceNames(self) -> Tuple[str]:
        """
        Provides the supported service names of the implementation, including also indirect service names.
        """
        return ("com.sun.star.ui.ActionTriggerContainer",)

    # endregion BaseServiceInfo

    # region XMultiServiceFactory
    def createInstance(self, name: str):
        self._logger.warning("createInstance is not implement in this context")
        return None

    def createInstanceWithArguments(self, name: str, args: Any):
        self._logger.warning("createInstanceWithArguments is not implement in this context")
        return None

    def getAvailableServiceNames(self):
        return ()

    # endregion XMultiServiceFactory

    # region XIndexContainer

    def hasElements(self) -> bool:
        return len(self._items) != 0

    def getElementType(self) -> Any:
        return uno.getTypeByName("com.sun.star.beans.XPropertyValue")

    def getCount(self):
        return len(self._items)

    def getByIndex(self, n) -> Any:
        if 0 <= n < len(self._items):
            return self._items[n]
        return None  # should be raise IllegalArgumentException

    def replaceByIndex(self, n, item) -> None:
        if 0 <= n < len(self._items):
            self._items[n] = item

    def insertByIndex(self, n, item) -> None:
        if 0 <= n <= len(self._items):
            self._items.insert(n, item)

    def removeByIndex(self, n) -> None:
        if 0 <= n < len(self._items):
            self._items.pop(n)

    # endregion XIndexContainer

    # region Dunder methods
    def __len__(self):
        return len(self._items)

    def __getitem__(self, key):
        return self.getByIndex(key)

    def __setitem__(self, key, value):
        self.replaceByIndex(key, value)

    def __delitem__(self, key):
        self.removeByIndex(key)

    def __iter__(self):
        return iter(self._items)

    # endregion Dunder methods

    # region Other Methods

    def append(self, item: Any):
        self._items.append(item)

    # endregion Other Methods
