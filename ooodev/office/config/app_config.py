from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter.ui.the_module_ui_configuration_manager_supplier_comp import (
    TheModuleUIConfigurationManagerSupplierComp,
)
from ooodev.utils import props as mProps


class AppConfig:
    def __init__(self):
        self._config_supp = TheModuleUIConfigurationManagerSupplierComp.from_lo()
        self._update_access = self._config_supp.get_ui_configuration_manager()

    def set_app_config(self, node_name: str, **kwargs) -> Any:
        update = self._config_supp.get_ui_configuration_manager(node_name)
        node_property = mProps.Props.make_prop_value(name="nodepath", value=node_name)
        for key, value in kwargs.items():
            update.setPropertyValue(key, value)
