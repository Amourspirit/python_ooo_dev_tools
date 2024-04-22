from __future__ import annotations
import pytest
from typing import cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.inst.lo_inst import LoInst

from ooodev.adapter.ui import global_accelerator_configuration_comp


def test_comp(loader, fix_lo_inst) -> None:
    lo_inst = cast(LoInst, fix_lo_inst)
    inst = global_accelerator_configuration_comp.GlobalAcceleratorConfigurationComp.from_lo(lo_inst=lo_inst)
    assert inst is not None
