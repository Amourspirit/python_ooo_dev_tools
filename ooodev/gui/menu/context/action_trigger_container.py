from __future__ import annotations
from typing import Any, Tuple

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

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
    @override
    def getImplementationName(self) -> str:
        """
        Provides the implementation name of the service implementation.
        """
        return "action_trigger_container"

    @override
    def supportsService(self, ServiceName: str) -> bool:
        """
        Tests whether the specified service is supported, i.e.

        implemented by the implementation.
        """
        return ServiceName in self.getAvailableServiceNames()

    @override
    def getSupportedServiceNames(self) -> Tuple[str]:
        """
        Provides the supported service names of the implementation, including also indirect service names.
        """
        return ("com.sun.star.ui.ActionTriggerContainer",)

    # endregion BaseServiceInfo

    # region XMultiServiceFactory
    @override
    def createInstance(self, aServiceSpecifier: str):
        """Raises NotImplementedError"""
        self._logger.warning("createInstance is not implement in this context")
        raise NotImplementedError

    @override
    def createInstanceWithArguments(self, ServiceSpecifier: str, Arguments: Any):
        """Raises NotImplementedError"""
        self._logger.warning("createInstanceWithArguments is not implement in this context")
        raise NotImplementedError

    def getAvailableServiceNames(self):
        return ()

    # endregion XMultiServiceFactory

    # region XIndexContainer

    @override
    def hasElements(self) -> bool:
        return len(self._items) != 0

    @override
    def getElementType(self) -> Any:
        return uno.getTypeByName("com.sun.star.beans.XPropertyValue")

    @override
    def getCount(self) -> int:
        return len(self._items)

    @override
    def getByIndex(self, Index: int) -> Any:
        if 0 <= Index < len(self._items):
            return self._items[Index]
        return None  # should be raise IllegalArgumentException

    @override
    def replaceByIndex(self, Index: int, Element: Any) -> None:
        if 0 <= Index < len(self._items):
            self._items[Index] = Element

    @override
    def insertByIndex(self, Index: int, Element: Any) -> None:
        if 0 <= Index <= len(self._items):
            self._items.insert(Index, Element)

    @override
    def removeByIndex(self, Index: int) -> None:
        if 0 <= Index < len(self._items):
            self._items.pop(Index)

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
