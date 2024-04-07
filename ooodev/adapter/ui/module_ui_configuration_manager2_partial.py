from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.ui import XModuleUIConfigurationManager2
from ooodev.adapter.ui.ui_configuration_persistence_partial import UIConfigurationPersistencePartial
from ooodev.adapter.ui.ui_configuration_manager_partial import UIConfigurationManagerPartial
from ooodev.adapter.ui.module_ui_configuration_manager_partial import ModuleUIConfigurationManagerPartial
from ooodev.adapter.ui.ui_configuration_partial import UIConfigurationPartial

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ModuleUIConfigurationManager2Partial(
    UIConfigurationPersistencePartial,
    UIConfigurationManagerPartial,
    ModuleUIConfigurationManagerPartial,
    UIConfigurationPartial,
):
    """
    Partial Class for XModuleUIConfigurationManager2.
    """

    def __init__(
        self,
        component: XModuleUIConfigurationManager2,
        interface: UnoInterface | None = XModuleUIConfigurationManager2,
    ) -> None:
        """
        Constructor

        Args:
            component (XModuleUIConfigurationManager2): UNO Component that implements ``com.sun.star.ui.XModuleUIConfigurationManager2``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XModuleUIConfigurationManager2``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        UIConfigurationPersistencePartial.__init__(self, component=component, interface=None)
        UIConfigurationManagerPartial.__init__(self, component=component, interface=None)
        ModuleUIConfigurationManagerPartial.__init__(self, component=component, interface=None)
        UIConfigurationPartial.__init__(self, component=component, interface=None)
