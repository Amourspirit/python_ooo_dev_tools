from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.ui import XUIConfigurationPersistence

from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.embed import XStorage
    from ooodev.utils.type_var import UnoInterface


class UIConfigurationPersistencePartial:
    """
    Partial Class for XUIConfigurationPersistence.
    """

    def __init__(
        self, component: XUIConfigurationPersistence, interface: UnoInterface | None = XUIConfigurationPersistence
    ) -> None:
        """
        Constructor

        Args:
            component (XUIConfigurationPersistence): UNO Component that implements ``com.sun.star.ui.XUIConfigurationPersistence``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUIConfigurationPersistence``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XUIConfigurationPersistence
    def is_modified(self) -> bool:
        """
        Provides the current modify state of the configuration manager instance.
        """
        return self.__component.isModified()

    def is_read_only(self) -> bool:
        """
        Provides the current read-only state of the user configuration manager.

        Storing a user interface configuration to a read-only storage is not possible.
        A read-only configuration manager instance will also not support any changes to its configuration settings.
        """
        return self.__component.isReadOnly()

    def reload(self) -> None:
        """
        Reloads the configuration data from the storage and reinitialize the user interface configuration manager instance with this data.

        It is up to the implementation if it defers the first loading process until the first data request using ``XUIConfigurationManager`` interface.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.reload()

    def store(self) -> None:
        """
        Stores the configuration data to the storage provided by ``setStorage()``
        from the storage and initialize the user interface configuration manager instance with the newly data.

        This call can throw an ``com.sun.star.io.IOException`` if ``store()`` cannot store its data into the internal storage.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.store()

    def store_to_storage(self, storage: XStorage) -> None:
        """
        stores the configuration data to the provided storage, ignoring the previously set storage by`` setStorage()``.

        Can be used to make copy of the current user interface configuration data to another storage.
        This call will throw an com.sun.star.io.IOException if the provided storage is in read-only mode.

        Raises:
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.storeToStorage(storage)

    # endregion XUIConfigurationPersistence


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)

    builder.auto_add_interface("com.sun.star.ui.XUIConfigurationPersistence")
    return builder
