from __future__ import annotations
import pytest
from typing import cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.inst.lo_inst import LoInst
from ooodev.adapter.deployment import extension_manager_comp


def test_comp(loader, fix_lo_inst) -> None:
    lo_inst = cast(LoInst, fix_lo_inst)
    inst = extension_manager_comp.ExtensionManagerComp.from_lo(lo_inst=lo_inst)
    assert inst is not None
    key = "com.sun.star.deployment.ExtensionManager"
    assert key in lo_inst.cache
    # test cached version
    inst = extension_manager_comp.ExtensionManagerComp.from_lo(lo_inst=lo_inst)
    assert inst is not None
