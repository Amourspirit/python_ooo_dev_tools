from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.ui import XUIConfigurationStorage

from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.embed import XStorage
    from ooodev.utils.type_var import UnoInterface


class UIConfigurationStoragePartial:
    """
    Partial Class for XUIConfigurationStorage.
    """

    def __init__(
        self, component: XUIConfigurationStorage, interface: UnoInterface | None = XUIConfigurationStorage
    ) -> None:
        """
        Constructor

        Args:
            component (XUIConfigurationStorage): UNO Component that implements ``com.sun.star.ui.XUIConfigurationStorage``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUIConfigurationStorage``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XUIConfigurationStorage
    def has_storage(self) -> bool:
        """
        Checks if an instance has already a storage to load/store its data.
        """
        return self.__component.hasStorage()

    def set_storage(self, storage: XStorage) -> None:
        """
        Connects a storage to the user interface configuration manager which is used on subsequent calls of load() and store().

        This call needs careful usage as data loss can occur. If the implementation has modified data and a new storage is set the implementation is not allowed to write back the data to the old storage. This must be done explicitly with store(). In general a user interface configuration manager instance is created and initialize by the document model or the module user interface configuration supplier. Normally there is no need to set a storage for a user interface configuration manager. If a copy of the configuration data to another storage should be made, use storeToStorage().
        """
        self.__component.setStorage(storage)

    # endregion XUIConfigurationStorage


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)

    builder.auto_add_interface("com.sun.star.ui.XUIConfigurationStorage")
    return builder
