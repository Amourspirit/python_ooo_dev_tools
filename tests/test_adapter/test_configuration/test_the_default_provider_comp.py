from __future__ import annotations
import pytest
from typing import cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.inst.lo_inst import LoInst
from ooodev.adapter.configuration import the_default_provider_comp


class _Dummy:
    pass


def test_comp(loader, fix_lo_inst) -> None:
    lo_inst = cast(LoInst, fix_lo_inst)
    inst = the_default_provider_comp.TheDefaultProviderComp.from_lo(lo_inst=lo_inst)
    assert inst is not None
    key = "com.sun.star.configuration.theDefaultProvider"
    assert key in lo_inst.cache
    # test cached version
    inst = the_default_provider_comp.TheDefaultProviderComp.from_lo(lo_inst=lo_inst)
    assert inst is not None


def test_builder(loader, fix_lo_inst):
    lo_inst = cast(LoInst, fix_lo_inst)
    inst = the_default_provider_comp.TheDefaultProviderComp.from_lo(lo_inst=lo_inst)
    builder = the_default_provider_comp.get_builder(inst.component)
    assert builder is not None
    clz = builder.build_class(name="Test", base_class=_Dummy)
    assert clz is not None
