from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
from unittest import result

from com.sun.star.ui import XModuleUIConfigurationManager

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.container.index_access_comp import IndexAccessComp

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ModuleUIConfigurationManagerPartial:
    """
    Partial Class for XModuleUIConfigurationManager.
    """

    def __init__(
        self, component: XModuleUIConfigurationManager, interface: UnoInterface | None = XModuleUIConfigurationManager
    ) -> None:
        """
        Constructor

        Args:
            component (XModuleUIConfigurationManager): UNO Component that implements ``com.sun.star.ui.XModuleUIConfigurationManager``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XModuleUIConfigurationManager``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XModuleUIConfigurationManager
    def get_default_settings(self, resource_url: str) -> IndexAccessComp:
        """
        Gets the settings from the default layer of the user interface configuration manager if it has a default layer.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        result = self.__component.getDefaultSettings(resource_url)
        if result is None:
            return None  # type: ignore
        return IndexAccessComp(result)

    def is_default_settings(self, resource_url: str) -> bool:
        """
        determine if the settings of a user interface element is part of the default layer of the user interface configuration manager.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        return self.__component.isDefaultSettings(resource_url)

    # endregion XModuleUIConfigurationManager
