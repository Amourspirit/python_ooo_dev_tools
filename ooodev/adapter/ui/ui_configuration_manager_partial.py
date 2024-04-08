from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

from com.sun.star.ui import XUIConfigurationManager

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.adapter.container.index_container_comp import IndexContainerComp
from ooodev.adapter.ui.accelerator_configuration_comp import AcceleratorConfigurationComp

if TYPE_CHECKING:
    from com.sun.star.container import XIndexAccess
    from com.sun.star.uno import XInterface
    from com.sun.star.beans import PropertyValue  # struct
    from ooodev.utils.type_var import UnoInterface


class UIConfigurationManagerPartial:
    """
    Partial Class for XUIConfigurationManager.
    """

    def __init__(
        self, component: XUIConfigurationManager, interface: UnoInterface | None = XUIConfigurationManager
    ) -> None:
        """
        Constructor

        Args:
            component (XUIConfigurationManager): UNO Component that implements ``com.sun.star.ui.XUIConfigurationManager``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUIConfigurationManager``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XUIConfigurationManager
    def create_settings(self) -> IndexContainerComp[Any]:
        """
        Creates an empty settings data container.
        """
        result = self.__component.createSettings()
        if result is None:
            return None  # type: ignore
        return IndexContainerComp(result)

    def get_events_manager(self) -> XInterface:
        """
        retrieves the events manager from the user interface configuration manager.

        Every user interface configuration manager has one events manager instance which controls the mapping of events to script URLs of a module or document.
        """
        return self.__component.getEventsManager()

    def get_image_manager(self) -> XInterface:
        """
        retrieves the image manager from the user interface configuration manager.

        Every user interface configuration manager has one image manager instance which controls all images of a module or document.
        """
        return self.__component.getImageManager()

    def get_settings(self, resource_url: str, writeable: bool) -> IndexAccessComp[Any]:
        """
        retrieves the settings of a user interface element.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        result = self.__component.getSettings(resource_url, writeable)
        if result is None:
            return None  # type: ignore
        return IndexAccessComp(result)

    def get_short_cut_manager(self) -> AcceleratorConfigurationComp:
        """
        Gets the keyboard short cut manager from the user interface configuration manager.

        Every user interface configuration manager has one keyboard short cut manager instance which controls all short cuts of a module or document.
        """
        result = self.__component.getShortCutManager()
        if result is None:
            return None  # type: ignore
        return AcceleratorConfigurationComp(result)

    def get_ui_elements_info(self, element_type: int) -> Tuple[Tuple[PropertyValue, ...], ...]:
        """
        Gets information about all user interface elements within the user interface configuration manager.

        The following ``PropertyValue`` entries are defined inside the sequence for every user interface element.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.getUIElementsInfo(element_type)

    def has_settings(self, resource_url: str) -> bool:
        """
        determines if the settings of a user interface element is part the user interface configuration manager.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.hasSettings(resource_url)

    def insert_settings(self, new_resource_url: str, new_data: XIndexAccess) -> None:
        """
        Insert the settings of a new user interface element.

        If the settings data is already present a ``com.sun.star.container.ElementExistException`` is thrown.
        If the ``new_resource_url`` is not valid or describes an unknown type a ``com.sun.star.lang.IllegalArgumentException`` is thrown.
        If the configuration manager is read-only a ``com.sun.star.lang.IllegalAccessException`` is thrown.

        Raises:
            com.sun.star.container.ElementExistException: ``ElementExistException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.IllegalAccessException: ``IllegalAccessException``
        """
        self.__component.insertSettings(new_resource_url, new_data)

    def remove_settings(self, ResourceURL: str) -> None:
        """
        removes the settings of an existing user interface element.

        If the settings data cannot be found a com.sun.star.container.NoSuchElementException is thrown. If the ResourceURL is not valid or describes an unknown type a com.sun.star.lang.IllegalArgumentException is thrown. If the configuration manager is read-only a com.sun.star.lang.IllegalAccessException is thrown.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.IllegalAccessException: ``IllegalAccessException``
        """
        ...

    def replace_settings(self, resource_url: str, new_data: XIndexAccess) -> None:
        """
        Replaces the settings of a user interface element with new settings.

        If the settings data cannot be found a ``com.sun.star.container.NoSuchElementException`` is thrown.
        If the ``resource_url`` is not valid or describes an unknown type a ``com.sun.star.lang.IllegalArgumentException`` is thrown.
        If the configuration manager is read-only a ``com.sun.star.lang.IllegalAccessException`` is thrown.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.IllegalAccessException: ``IllegalAccessException``
        """
        self.__component.replaceSettings(resource_url, new_data)

    def reset(self) -> None:
        """
        Resets the configuration manager to the default user interface configuration data.

        This means that all user interface configuration data of the instance will be removed.
        A module based user interface configuration manager removes user defined elements, but set all other elements back to default.
        It is not possible to remove default elements from a module user interface configuration manager.
        """
        self.__component.reset()

    # endregion XUIConfigurationManager
