from __future__ import annotations
import pytest
from typing import cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.inst.lo_inst import LoInst

from ooodev.adapter.ui import the_module_ui_configuration_manager_supplier_comp


def test_comp(loader, fix_lo_inst) -> None:
    lo_inst = cast(LoInst, fix_lo_inst)
    inst = the_module_ui_configuration_manager_supplier_comp.TheModuleUIConfigurationManagerSupplierComp.from_lo(
        lo_inst=lo_inst
    )
    assert inst is not None
    key = "com.sun.star.ui.theModuleUIConfigurationManagerSupplier"
    assert key in lo_inst.cache
    # test cached version
    inst = the_module_ui_configuration_manager_supplier_comp.TheModuleUIConfigurationManagerSupplierComp.from_lo(
        lo_inst=lo_inst
    )
    assert inst is not None
