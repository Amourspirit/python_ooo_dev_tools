from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooo.dyn.awt.key import KeyEnum
from ooo.dyn.awt.key_event import KeyEvent
from ooo.dyn.awt.key_modifier import KeyModifierEnum
from ooo.dyn.beans.property_concept import PropertyConceptEnum

from ooodev.loader.inst.service import Service
from ooodev.adapter.ui.the_module_ui_configuration_manager_supplier_comp import (
    TheModuleUIConfigurationManagerSupplierComp,
)

if TYPE_CHECKING:
    from ooodev.adapter.ui.ui_configuration_manager_comp import UIConfigurationManagerComp


class MenuApp:
    def __init__(self, app: Service) -> None:
        self._app = app
        self._config = self._get_config()

    def _get_config(self) -> UIConfigurationManagerComp:
        supp = TheModuleUIConfigurationManagerSupplierComp.from_lo()
        config = supp.get_ui_configuration_manager(self._app)
        return config
