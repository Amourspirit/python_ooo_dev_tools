from __future__ import annotations
from logging import info
from typing import Any, TYPE_CHECKING

from com.sun.star.ui import XModuleUIConfigurationManagerSupplier

from ooodev.adapter.ui.ui_configuration_manager_comp import UIConfigurationManagerComp
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from ooodev.loader.inst.service import Service


class ModuleUIConfigurationManagerSupplierPartial:
    """
    Partial Class for XModuleUIConfigurationManagerSupplier.
    """

    def __init__(
        self,
        component: XModuleUIConfigurationManagerSupplier,
        interface: UnoInterface | None = XModuleUIConfigurationManagerSupplier,
    ) -> None:
        """
        Constructor

        Args:
            component (XModuleUIConfigurationManagerSupplier): UNO Component that implements ``com.sun.star.ui.XModuleUIConfigurationManagerSupplier``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XModuleUIConfigurationManagerSupplier``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    def __get_configuration_manager(self, module_identifier: str | Service) -> Any:
        """
        Gets the requested module based user interface configuration manager.

        Args:
            module_identifier (str): The identifier of the module for which the user interface configuration manager is requested.
                Such as ``com.sun.star.sheet.SpreadsheetDocument``.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``

        Returns:
            UIConfigurationManagerComp: The requested module based user interface configuration manager.

        Hint:
            - ``Service`` is an enum and can be imported from ``ooodev.loader.inst.service.Service``.
        """
        # 'com.sun.star.comp.framework.ModuleUIConfigurationManager'
        mgr = self.__component.getUIConfigurationManager(str(module_identifier))
        if mgr is None:
            return None
        name = mInfo.Info.get_implementation_name(mgr)
        if name == "com.sun.star.comp.framework.ModuleUIConfigurationManager":
            from ooodev.adapter.ui.module_ui_configuration_manager_comp import ModuleUIConfigurationManagerComp

            return ModuleUIConfigurationManagerComp(mgr)  # type: ignore
        return UIConfigurationManagerComp(mgr)

    # region XModuleUIConfigurationManagerSupplier
    def get_ui_configuration_manager(self, module_identifier: str | Service) -> UIConfigurationManagerComp:
        """
        Gets the requested module based user interface configuration manager.

        Args:
            module_identifier (str): The identifier of the module for which the user interface configuration manager is requested.
                Such as ``com.sun.star.sheet.SpreadsheetDocument``.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``

        Returns:
            UIConfigurationManagerComp: The requested module based user interface configuration manager.

        Hint:
            - ``Service`` is an enum and can be imported from ``ooodev.loader.inst.service.Service``.
        """
        return self.__get_configuration_manager(module_identifier)  # type: ignore

    # endregion XModuleUIConfigurationManagerSupplier
